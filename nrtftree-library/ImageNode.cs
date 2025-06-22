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
 * Class:		ImageNode
 * Description:	Specialized RTF node that contains the information of an image.
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;
using System.Globalization;
using SkiaSharp;

namespace Net.Sgoliver.NRtfTree.Util;

public enum ImageFormat
{
    Unknown = 0,
    Jpeg = 1,
    Png = 2,
    Emf = 3,
    Wmf = 4,
    Bmp = 5,
}

/// <summary>
/// Encapsulates an RTF node of type Image (Keyword "\pict")
/// </summary>
public class ImageNode : Net.Sgoliver.NRtfTree.Core.RtfTreeNode
{

    /// <summary>
    /// Byte array with the image information.
    /// </summary>
    private byte[]? data;

    #region Constructor

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="node">RTF node from which the image data will be obtained.</param>
    public ImageNode(RtfTreeNode? node)
    {
        if (node == null) return;
        
        // We assign all the fields of the node
        NodeKey = node.NodeKey;
        HasParameter = node.HasParameter;
        Parameter = node.Parameter;
        ParentNode = node.ParentNode;
        RootNode = node.RootNode;
        NodeType = node.NodeType;

        ChildNodes = new RtfNodeCollection();
        if(node.ChildNodes != null)
            ChildNodes.AddRange(node.ChildNodes);

        // We get the image data as a byte array
        getImageData();
    }

    #endregion

    #region Properties

	/// <summary>
	/// Returns a string containing the image content in hexadecimal format.
	/// </summary>
	public string? HexData => SelectSingleChildNode(RtfNodeType.Text)?.NodeKey;

    /// <summary>
    /// Returns the original format of the image.
    /// </summary>
    public ImageFormat ImageFormat
    { 
        get 
        {
            if (SelectSingleChildNode("jpegblip") != null)
                return ImageFormat.Jpeg;
            else if (SelectSingleChildNode("pngblip") != null)
                return ImageFormat.Png;
            else if (SelectSingleChildNode("emfblip") != null)
                return ImageFormat.Emf;
            else if (SelectSingleChildNode("wmetafile") != null)
                return ImageFormat.Wmf;
            else if (SelectSingleChildNode("dibitmap") != null || SelectSingleChildNode("wbitmap") != null)
                return ImageFormat.Bmp;
            else
                return ImageFormat.Unknown;
        }
    }

    /// <summary>
    /// Returns the width of the image (in twips).
    /// </summary>
    public int Width
    {
        get
        {
            var node = SelectSingleChildNode("picw");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the height of the image (in twips).
    /// </summary> 
    public int Height
    {
        get
        {
            var node = SelectSingleChildNode("pich");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the target width of the image (in twips).
    /// </summary>
    public int DesiredWidth
    {
        get
        {
            var node = SelectSingleChildNode("picwgoal");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the target height of the image (in twips).
    /// </summary>
    public int DesiredHeight
    {
        get
        {
            var node = SelectSingleChildNode("pichgoal");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the horizontal scale of the image, in percentage.
    /// </summary>
    public int ScaleX
    {
        get
        {
            var node = SelectSingleChildNode("picscalex");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the vertical scale of the image, in percentage.
    /// </summary>
    public int ScaleY
    {
        get
        {
            var node = SelectSingleChildNode("picscaley");

            if (node != null)
                return node.Parameter;
            else
                return -1;
        }
    }

    /// <summary>
    /// Returns the image in a bitmap object.
    /// </summary>
    public SKBitmap Bitmap => data == null
        ? new SKBitmap()
        : SKBitmap.Decode(data);

    #endregion

    #region Public Methods

	/// <summary>
	/// Returns a byte array containing the contents of the image.
	/// </summary>
	/// <return>Byte array containing the image content.</return>
	public byte[]? GetByteData()
	{
		return data;
	}

    /// <summary>
    /// Save an image to a file in its original format.
    /// </summary>
    /// <param name="filePath">Path of the file where the image will be saved.</param>
    public void SaveImage(string filePath)
    {
        if (data == null) return;
        
        var stream = File.OpenWrite(filePath);

        // Write any type of image to a file
        switch (ImageFormat)
        {
            case ImageFormat.Jpeg: Bitmap.Encode(SKEncodedImageFormat.Jpeg, 90).SaveTo(stream); break;
            case ImageFormat.Emf: /* unsupported */ break;
            case ImageFormat.Wmf: /* unsupported */ break;
            case ImageFormat.Bmp: Bitmap.Encode(SKEncodedImageFormat.Bmp, 90).SaveTo(stream); break;
            default:
            case ImageFormat.Png: Bitmap.Encode(SKEncodedImageFormat.Png, 90).SaveTo(stream); break;
        }

        stream.Flush();
        stream.Close();
    }

    /// <summary>
    /// Saves an image to a file with a specific format specified as a parameter.
    /// </summary>
    /// <param name="filePath">Path of the file where the image will be saved.</param>
    /// <param name="format">Format in which the image will be written.</param>
    public void SaveImage(string filePath, ImageFormat ImageFormat)
    {
        if (data == null) return;
        var stream = File.OpenWrite(filePath);

        // Write any type of image to a file
        switch (ImageFormat)
        {
            case ImageFormat.Jpeg: Bitmap.Encode(SKEncodedImageFormat.Jpeg, 90).SaveTo(stream); break;
            case ImageFormat.Emf: /* unsupported */ break;
            case ImageFormat.Wmf: /* unsupported */ break;
            case ImageFormat.Bmp: Bitmap.Encode(SKEncodedImageFormat.Bmp, 90).SaveTo(stream); break;
            default:
            case ImageFormat.Png: Bitmap.Encode(SKEncodedImageFormat.Png, 90).SaveTo(stream); break;
        }

        stream.Flush();
        stream.Close();
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Obtains image data from the information contained in the RTF node.
    /// </summary>
    private void getImageData()
    {
        //Format 1 (Word 97-2000): {\*\shppict {\pict\jpegblip <datos>}}{\nonshppict {\pict\wmetafile8 <datos>}}
        //Format 2 (Wordpad)     : {\pict\wmetafile8 <datos>}

        var text = "";

        if (FirstChild?.NodeKey != "pict") return;
        
        text = SelectSingleChildNode(RtfNodeType.Text)?.NodeKey ?? "";

        var dataSize = text.Length / 2;
        data = new byte[dataSize];

        var sbaux = new StringBuilder(2);

        for (var i = 0; i < text.Length; i++)
        {
            sbaux.Append(text[i]);

            if (sbaux.Length != 2) continue;
            
            data[i / 2] = byte.Parse(sbaux.ToString(), NumberStyles.HexNumber);
            sbaux.Remove(0, 2);
        }
    }

    #endregion
}

