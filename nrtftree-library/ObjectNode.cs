/********************************************************************************
 *   This file is part of NRtfTree Library.
 *
 *   NRtfTree Library is free software; you can redistribute it and/or modify
 *   it under the terms of the GNU Lesser General Public License as published by
 *   the Free Software Foundation; either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   NRtfTree Library is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU Lesser General Public License for more details.
 *
 *   You should have received a copy of the GNU Lesser General Public License
 *   along with this program. If not, see <http://www.gnu.org/licenses/>.
 ********************************************************************************/

/********************************************************************************
 * Library:		NRtfTree
 * Version:     v0.4
 * Date:		29/06/2013
 * Copyright:   2006-2013 Salvador Gomez
 * Home Page:	http://www.sgoliver.net
 * GitHub:	    https://github.com/sgolivernet/nrtftree
 * Class:		ObjectNode
 * Description:	Specialized RTF node that contains an object's information.
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;
using System.Globalization;

namespace Net.Sgoliver.NRtfTree.Util;

/// <summary>
/// Encapsulates an RTF node of type Object (Keyword "\object")
/// </summary>
public class ObjectNode : Net.Sgoliver.NRtfTree.Core.RtfTreeNode
{
    /// <summary>
    /// Byte array containing the object information.
    /// </summary>
    private byte[]? objdata;

    #region Constructors

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="node">RTF node from which the image data will be obtained.</param>
    public ObjectNode(RtfTreeNode? node)
    {
        if (node == null) return;
        
        // We assign all the fields of the node
        NodeKey = node.NodeKey;
        HasParameter = node.HasParameter;
        Parameter= node.Parameter;
        ParentNode = node.ParentNode;
        RootNode = node.RootNode;
        NodeType = node.NodeType;

        ChildNodes = new RtfNodeCollection();
        if (node.ChildNodes != null)
            ChildNodes.AddRange(node.ChildNodes);

        // We get the object data as a byte array
        getObjectData();
    }

    #endregion

    #region Properties

    /// <summary>
    /// Returns the type of the object.
    /// </summary>
    public string ObjectType
    {
        get 
        {
            if (SelectSingleChildNode("objemb") != null)
                return "objemb";
            if (SelectSingleChildNode("objlink") != null)
                return "objlink";
            if (SelectSingleChildNode("objautlink") != null)
                return "objautlink";
            if (SelectSingleChildNode("objsub") != null)
                return "objsub";
            if (SelectSingleChildNode("objpub") != null)
                return "objpub";
            if (SelectSingleChildNode("objicemb") != null)
                return "objicemb";
            if (SelectSingleChildNode("objhtml") != null)
                return "objhtml";
            if (SelectSingleChildNode("objocx") != null)
                return "objocx";
            else
                return "";
        }
    }

    /// <summary>
    /// Returns the type of the object.
    /// </summary>
    public string? ObjectClass
    {
        get
        {
            //Format: {\*\objclass Paint.Picture}

            var node = SelectSingleNode("objclass");

            return node != null ? node.NextSibling?.NodeKey : "";
        }
    }

    /// <summary>
    /// Returns the RTF group that encapsulates the "\result" node of the object.
    /// </summary>
    public RtfTreeNode? ResultNode
    {
        get
        {
            var node = SelectSingleNode("result");

            // If the "\result" node exists, we retrieve the top RTF group.
            node = node?.ParentNode;

            return node;
        }
    }

	/// <summary>
	/// Returns a string containing the contents of the object in hexadecimal format.
	/// </summary>
	public string HexData
	{
		get
		{
			var text = "";

			// We look for the "\objdata" node
			var objdataNode = SelectSingleNode("objdata");

			// If the node exists
			if (objdataNode != null)
			{
				// We look for the data in hexadecimal format (last child of the \objdata group)
				text = objdataNode.ParentNode?.LastChild?.NodeKey ?? "";
			}

			return text;				
		}
	}

    #endregion

	#region Public Methods

	/// <summary>
	/// Returns a byte array containing the contents of the object.
	/// </summary>
	/// <returns>Byte array containing the contents of the object.</returns>
	public byte[]? GetByteData()
	{
		return objdata;
	}

	#endregion

    #region Private Methods

    /// <summary>
    /// Obtains the binary data of the object from the information contained in the RTF node.
    /// </summary>
    private void getObjectData()
    {
        //Format: ( '{' \object (<objtype> & <objmod>? & <objclass>? & <objname>? & <objtime>? & <objsize>? & <rsltmod>?) ('{\*' \objdata (<objalias>? & <objsect>?) <data> '}') <result> '}' )

        var text = "";

        if (FirstChild?.NodeKey != "object") return;
        
        // We look for the "\objdata" node
        var objdataNode = SelectSingleNode("objdata");

        if (objdataNode?.ParentNode?.LastChild == null) return;
        
        // We look for the data in hexadecimal format (last child of the \objdata group)
        text = objdataNode!.ParentNode!.LastChild!.NodeKey;

        var dataSize = text.Length / 2;
        objdata = new byte[dataSize];

        var sbaux = new StringBuilder(2);

        for (var i = 0; i < text.Length; i++)
        {
            sbaux.Append(text[i]);

            if (sbaux.Length != 2) continue;
            
            objdata[i / 2] = byte.Parse(sbaux.ToString(), NumberStyles.HexNumber);
            sbaux.Remove(0, 2);
        }
    }

    #endregion
}
