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
 * Class:		RtfTreeNode
 * Description:	RTF node of the tree representation of a document.
 * ******************************************************************************/

using System.Text;

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// RTF node of the tree representation of a document.
/// </summary>
public class RtfTreeNode
{
    #region Private Fields

    /// <summary>
    /// Child nodes of the current node.
    /// </summary>
    private RtfNodeCollection? children;

    #endregion

    #region Constructors

    /// <summary>
    /// Constructor for the RtfTreeNode class. Creates an uninitialized node.
    /// </summary>
    public RtfTreeNode()
    {
        NodeType = RtfNodeType.None;
        NodeKey = "";
    }

    /// <summary>
    /// Constructor for the RtfTreeNode class. Creates a node of a specific type.
    /// </summary>
    /// <param name="nodeType">Type of node to create.</param>
    public RtfTreeNode(RtfNodeType nodeType)
    {
        NodeType = nodeType;
        NodeKey = "";

        if (nodeType is RtfNodeType.Group or RtfNodeType.Root)
            children = new RtfNodeCollection();

        if (nodeType == RtfNodeType.Root)
            RootNode = this;

    }

    /// <summary>
    /// Constructor for the RtfTreeNode class. Creates a node by specifying its type, keyword, and parameter.
    /// </summary>
    /// <param name="type">Type of node.</param>
    /// <param name="key">Keyword or Control symbol.</param>
    /// <param name="hasParameter">Indicates whether the keyword or control symbol is accompanied by a parameter.</param>
    /// <param name="parameter">Parameter of the Control keyword or symbol.</param>
    public RtfTreeNode(RtfNodeType type, string key, bool hasParameter, int parameter)
    {
        this.NodeType = type;
        this.NodeKey = key;
        HasParameter = hasParameter;
        Parameter = parameter;

        if (type is RtfNodeType.Group or RtfNodeType.Root)
            children = new RtfNodeCollection();

        if (type == RtfNodeType.Root)
            RootNode = this;

    }

    #endregion

    #region Private Constructor

    /// <summary>
    /// Private constructor for the RtfTreeNode class. Creates a node from a lexical analyzer token.
    /// </summary>
    /// <param name="token">RTF token returned by the lexical analyzer.</param>
    internal RtfTreeNode(RtfToken token)
    {
        NodeType = (RtfNodeType)token.Type;
        NodeKey = token.Key;
        HasParameter = token.HasParameter;
        Parameter = token.Parameter;
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Adds a node to the end of the children list.
    /// </summary>
    /// <param name="newNode">New node to add.</param>
    public void AppendChild(RtfTreeNode? newNode)
    {
        if (newNode == null) return;
        // If you do not have children yet, the collection is initialized.
        children ??= new RtfNodeCollection();

        // The current node is assigned as the parent node
        newNode.ParentNode = this;

        // The Root and Tree properties of the new node and its possible children are updated
        updateNodeRoot(newNode);

        // The new node is added to the end of the list of child nodes
        children.Add(newNode);
    }

    /// <summary>
    /// Inserts a new node at a given position in the list of children.
    /// </summary>
    /// <param name="index">Position at which the node will be inserted.</param>
    /// <param name="newNode">New node to add.</param>
    public void InsertChild(int index, RtfTreeNode? newNode)
    {
        if (newNode == null) return;
        // If you do not have children yet, the collection is initialized.
        children ??= new RtfNodeCollection();

        if (index < 0 || index > children.Count) return;
        
        // The current node is assigned as the parent node
        newNode.ParentNode = this;

        // The Root and Tree properties of the new node and its possible children are updated
        updateNodeRoot(newNode);

        // The new node is added to the end of the list of child nodes
        children.Insert(index, newNode);
    }

    /// <summary>
    /// Removes a node from the list of children.
    /// </summary>
    /// <param name="index">Indice del nodo a eliminar.</param>
    public void RemoveChild(int index)
    {
        // If the current node has children
        if (children == null) return;
        if (index >= 0 && index < children.Count)
        {
            // The i-th child is eliminated
            children.RemoveAt(index);
        }
    }

    /// <summary>
    /// Removes a node from the list of children.
    /// </summary>
    /// <param name="node">Nodo a eliminar.</param>
    public void RemoveChild(RtfTreeNode node)
    {
        // If the current node has children
        if (children == null) return;
        
        // Search for the node to be deleted
        var index = children.IndexOf(node);

        // If we find it
        if (index != -1)
        {
            // The i-th child is eliminated
            children.RemoveAt(index);
        }
    }

    /// <summary>
    /// Makes an exact copy of the current node.
    /// </summary>
    /// <returns>Returns an exact copy of the current node</returns>
    public RtfTreeNode CloneNode()
    {
        var clon = new RtfTreeNode
        {
            NodeKey = NodeKey,
            HasParameter = HasParameter,
            Parameter = Parameter,
            ParentNode = null,
            RootNode = null,
            Tree = null,
            NodeType = NodeType,
            // Each of the children is also cloned
            children = null
        };

        if (children == null) return clon;
        
        clon.children = new RtfNodeCollection();

        foreach (RtfTreeNode child in children)
        {
            var childclon = child.CloneNode();
            childclon.ParentNode = clon;

            clon.children.Add(childclon);
        }

        return clon;
    }

    /// <summary>
    /// Indicates whether the current node has any child nodes.
    /// </summary>
    /// <returns>Returns true if the current node has any child nodes.</returns>
    public bool HasChildNodes()
    {
        return children is { Count: > 0 };
    }

    /// <summary>
    /// Returns the first node in the list of child nodes of the current node whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>First node in the list of child nodes of the current node whose keyword is the one given as a parameter.</returns>
    public RtfTreeNode? SelectSingleChildNode(string keyword)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeKey == keyword)
            {
                node = children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first node in the list of child nodes of the current node whose type is the one specified as a parameter.
    /// </summary>
    /// <param name="nodeType">Type of node searched</param>
    /// <returns>First node in the list of child nodes of the current node whose type is the one specified as a parameter.</returns>
    public RtfTreeNode? SelectSingleChildNode(RtfNodeType nodeType)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        while (i < children.Count && !found)
        {
            if (children[i] != null && children[i]!.NodeType == nodeType)
            {
                node = children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first node in the list of child nodes of the current node whose keyword and parameter are the given parameters.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="parameter">Parameter searched.</param>
    /// <returns>First node in the list of child nodes of the current node whose keyword and parameter are those indicated as parameters</returns>
    public RtfTreeNode? SelectSingleChildNode(string keyword, int parameter)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeKey == keyword && children[i]?.Parameter == parameter)
            {
                node = children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first group node from the list of child nodes of the current node whose first keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="ignoreSpecial">If enabled, '\*' control nodes before some keywords will be ignored.</param>
    /// <returns>First group node in the list of child nodes of the current node whose first keyword is the one given as a parameter.</returns>
    public RtfTreeNode? SelectSingleChildGroup(string keyword, bool ignoreSpecial = false)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeType == RtfNodeType.Group && children[i]?.HasChildNodes() == true &&
                (
                    (children[i]?.FirstChild?.NodeKey == keyword) ||
                    (ignoreSpecial && children[i]?.ChildNodes?[0]?.NodeKey == "*" && children[i]?.ChildNodes?[1]?.NodeKey == keyword))
               )
            {
                node = children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first node in the tree, starting from the current node, whose type is the one specified as a parameter.
    /// </summary>
    /// <param name="nodeType">Type of node searched.</param>
    /// <returns>First node in the tree, starting from the current node, whose type is the one indicated as a parameter.</returns>
    public RtfTreeNode? SelectSingleNode(RtfNodeType nodeType)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeType == nodeType)
            {
                node = children[i];
                found = true;
            }
            else
            {
                node = children[i]?.SelectSingleNode(nodeType);

                if (node != null)
                {
                    found = true;
                }
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first node in the tree, starting from the current node, whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>First node of the tree, starting from the current node, whose keyword is the one indicated as a parameter.</returns>
    public RtfTreeNode? SelectSingleNode(string keyword)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeKey == keyword)
            {
                node = children[i];
                found = true;
            }
            else
            {
                node = children[i]?.SelectSingleNode(keyword);

                if (node != null)
                {
                    found = true;
                }
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first group node in the tree, starting from the current node, whose first keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="ignoreSpecial">If enabled, '\*' control nodes before some keywords will be ignored.</param>
    /// <returns>First group node of the tree, starting from the current node, whose first keyword is the one indicated as a parameter</returns>
    public RtfTreeNode? SelectSingleGroup(string keyword, bool ignoreSpecial = false)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeType == RtfNodeType.Group && children[i]?.HasChildNodes() == true &&
                (
                    (children[i]?.FirstChild?.NodeKey == keyword) ||
                    (ignoreSpecial && children[i]?.ChildNodes?[0]?.NodeKey == "*" && children[i]?.ChildNodes?[1]?.NodeKey == keyword))
               )
            {
                node = children[i];
                found = true;
            }
            else
            {
                node = children[i]?.SelectSingleGroup(keyword, ignoreSpecial);

                if (node != null)
                {
                    found = true;
                }
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the first node in the tree, starting from the current node, whose keyword and parameter are the ones given as parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="parameter">Parameter searched.</param>
    /// <returns>First node of the tree, starting from the current node, whose keyword and parameter are the ones given as parameter.</returns>
    public RtfTreeNode? SelectSingleNode(string keyword, int parameter)
    {
        var i = 0;
        var found = false;
        RtfTreeNode? node = null;

        if (children == null) return node;
        
        while (i < children.Count && !found)
        {
            if (children[i]?.NodeKey == keyword && children[i]?.Parameter == parameter)
            {
                node = children[i];
                found = true;
            }
            else
            {
                node = children[i]?.SelectSingleNode(keyword, parameter);

                if (node != null)
                {
                    found = true;
                }
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns all nodes, starting from the current node, whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>Collection of nodes, starting from the current node, whose keyword is the one given as a parameter.</returns>
    public RtfNodeCollection SelectNodes(string keyword)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeKey == keyword)
            {
                nodes.Add(node);
            }

            nodes.AddRange(node.SelectNodes(keyword));
        }

        return nodes;
    }

    /// <summary>
    /// Returns all group nodes, starting from the current node, whose first keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="ignoreSpecial">If enabled, '\*' control nodes before some keywords will be ignored.</param>
    /// <returns>Collection of nodes, starting from the current node, whose keyword is the one given as a parameter.</returns>
    public RtfNodeCollection SelectGroups(string keyword, bool ignoreSpecial = false)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == RtfNodeType.Group && node.HasChildNodes() &&
                (
                    (node.FirstChild?.NodeKey == keyword) ||
                    (ignoreSpecial && node.ChildNodes?[0]?.NodeKey == "*" && node.ChildNodes[1]?.NodeKey == keyword))
               )
            {
                nodes.Add(node);
            }

            nodes.AddRange(node.SelectGroups(keyword, ignoreSpecial));
        }

        return nodes;
    }

    /// <summary>
    /// Returns all nodes, starting from the current node, whose type is the one specified as a parameter.
    /// </summary>
    /// <param name="nodeType">Type of node searched.</param>
    /// <returns>Collection of nodes, starting from the current node, whose type is the one indicated as a parameter.</returns>
    public RtfNodeCollection SelectNodes(RtfNodeType nodeType)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == nodeType)
            {
                nodes.Add(node);
            }

            nodes.AddRange(node.SelectNodes(nodeType));
        }

        return nodes;
    }

    /// <summary>
    /// Returns all nodes, starting from the current node, whose keyword and parameter are the given parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="parameter">Parameter searched.</param>
    /// <returns>Collection of nodes, starting from the current node, whose keyword and parameter are those indicated as parameter.</returns>
    public RtfNodeCollection SelectNodes(string keyword, int parameter)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeKey == keyword && node.Parameter == parameter)
            {
                nodes.Add(node);
            }

            nodes.AddRange(node.SelectNodes(keyword, parameter));
        }

        return nodes;
    }

    /// <summary>
    /// Returns all nodes in the list of child nodes of the current node whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>Collection of nodes from the list of child nodes of the current node whose keyword is the one given as a parameter.</returns>
    public RtfNodeCollection SelectChildNodes(string keyword)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeKey == keyword)
            {
                nodes.Add(node);
            }
        }

        return nodes;
    }

    /// <summary>
    /// Returns all group nodes from the list of child nodes of the current node whose first keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="ignoreSpecial">If enabled, '\*' control nodes before some keywords will be ignored.</param>
    /// <returns>Collection of nodes group the list of child nodes of the current node whose first keyword is the one given as a parameter.</returns>
    public RtfNodeCollection SelectChildGroups(string keyword, bool ignoreSpecial = false)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == RtfNodeType.Group && node.HasChildNodes() &&
                (
                    (node.FirstChild?.NodeKey == keyword) ||
                    (ignoreSpecial && node.ChildNodes?[0]?.NodeKey == "*" && node.ChildNodes[1]?.NodeKey == keyword))
               )
            {
                nodes.Add(node);
            }
        }

        return nodes;
    }

    /// <summary>
    /// Returns all nodes in the list of child nodes of the current node whose type is the one specified as a parameter.
    /// </summary>
    /// <param name="nodeType">Type of node searched.</param>
    /// <returns>Collection of nodes from the list of child nodes of the current node whose type is the one specified as a parameter.</returns>
    public RtfNodeCollection SelectChildNodes(RtfNodeType nodeType)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == nodeType)
            {
                nodes.Add(node);
            }
        }

        return nodes;
    }

    /// <summary>
    /// Returns all nodes in the child node list of the current node whose keyword and parameter are the given parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="parameter">Parameter searched.</param>
    /// <returns>Collection of nodes from the list of child nodes of the current node whose keyword and parameter are those given as parameter.</returns>
    public RtfNodeCollection SelectChildNodes(string keyword, int parameter)
    {
        var nodes = new RtfNodeCollection();

        if (children == null) return nodes;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeKey == keyword && node.Parameter == parameter)
            {
                nodes.Add(node);
            }
        }

        return nodes;
    }

    /// <summary>
    /// Returns the next sibling node of the current node whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>First sibling node of the current one whose keyword is the one given as a parameter.</returns>
    public RtfTreeNode? SelectSibling(string keyword)
    {
        RtfTreeNode? node = null;
        var par = ParentNode;

        if (par == null) return node;

        if (par.ChildNodes == null) return node;
        
        var curInd = par.ChildNodes.IndexOf(this);

        var i = curInd + 1;
        var found = false;

        while (par.children != null && i < par.children.Count && !found)
        {
            if (par.children[i]?.NodeKey == keyword)
            {
                node = par.children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the next sibling node of the current node whose type is the one specified as a parameter.
    /// </summary>
    /// <param name="nodeType">Type of node searched.</param>
    /// <returns>First sibling node of the current one whose type is the one indicated as a parameter.</returns>
    public RtfTreeNode? SelectSibling(RtfNodeType nodeType)
    {
        RtfTreeNode? node = null;
        var par = ParentNode;

        if (par == null) return node;
        if (par.ChildNodes == null) return node;
        
        var curInd = par.ChildNodes.IndexOf(this);

        var i = curInd + 1;
        var found = false;

        while (par.children != null && i < par.children.Count && !found)
        {
            if (par.children[i]?.NodeType == nodeType)
            {
                node = par.children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Returns the next sibling node of the current node whose keyword and parameter are the given parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <param name="parameter">Parameter searched.</param>
    /// <returns>First sibling node of the current one whose keyword and parameter are those indicated as parameter.</returns>
    public RtfTreeNode? SelectSibling(string keyword, int parameter)
    {
        RtfTreeNode? node = null;
        var par = ParentNode;

        if (par == null) return node;
        if (par.ChildNodes == null) return node;
        
        var curInd = par.ChildNodes.IndexOf(this);

        var i = curInd + 1;
        var found = false;

        while (par.children != null && i < par.children.Count && !found)
        {
            if (par.children[i]?.NodeKey == keyword && par.children[i]?.Parameter == parameter)
            {
                node = par.children[i];
                found = true;
            }

            i++;
        }

        return node;
    }

    /// <summary>
    /// Finds all nodes of type Text that contain the searched text.
    /// </summary>
    /// <param name="text">Text searched for in the document.</param>
    /// <returns>List of nodes, starting from the current one, that contain the search text.</returns>
    public RtfNodeCollection FindText(string text)
    {
        var list = new RtfNodeCollection();

        if (children == null) return list;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == RtfNodeType.Text && node.NodeKey.Contains(text))
                list.Add(node);
            else if(node.NodeType == RtfNodeType.Group)
                list.AddRange(node.FindText(text));
        }

        return list;
    }

    /// <summary>
    /// Finds and replaces a given text in all nodes of type Text starting from the current one.
    /// </summary>
    /// <param name="oldValue">Text searched for in the document.</param>
    /// <param name="newValue">Text to replace with.</param>
    public void ReplaceText(string oldValue, string newValue)
    {
        if (children == null) return;
        
        foreach (RtfTreeNode node in children)
        {
            if (node.NodeType == RtfNodeType.Text)
                node.NodeKey = node.NodeKey.Replace(oldValue, newValue);
            else if (node.NodeType == RtfNodeType.Group)
                node.ReplaceText(oldValue, newValue);
        }
    }

    /// <summary>
    /// Returns a representation of the node where its type, key, parameter indicator and parameter value are indicated
    /// </summary>
    /// <returns>Character string of type [TYPE, KEY, IND_PARAMETER, VAL_PARAMETER]</returns>
    public override string ToString()
    {
        return $"[{NodeType}, {NodeKey}, {HasParameter}, {Parameter}]";
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Decodes a special character indicated by its decimal code
    /// </summary>
    /// <param name="code">Special character code (\')</param>
    /// <param name="enc">Encoding used to decode the special character.</param>
    /// <returns>Decoded special character.</returns>
    private string DecodeControlChar(int code, Encoding enc)
    {
        //Contributed by Jan Stuchlík
        return enc.GetString(new[] { (byte)code });
    }

    /// <summary>
    /// Obtains the RTF Text from the tree representation of the current node.
    /// </summary>
    /// <returns>RTF text of the node.</returns>
    private string getRtf()
    {
        var enc = Tree?.GetEncoding();

        return getRtfInm(this, null, enc);
    }

    /// <summary>
    /// Helper method to obtain the RTF Text of the current node from its tree representation.
    /// </summary>
    /// <param name="curNode">Current node of the tree.</param>
    /// <param name="prevNode">Previous node.</param>
    /// <param name="enc">Document encoding.</param>
    /// <returns>Text in RTF format of the node.</returns>
    private string getRtfInm(RtfTreeNode curNode, RtfTreeNode? prevNode, Encoding? enc)
    {
        var res = new StringBuilder("");

        switch (curNode.NodeType)
        {
            case RtfNodeType.Root:
                res.Append("");
                break;
            case RtfNodeType.Group:
                res.Append('{');
                break;
            default:
            {
                if (curNode.NodeType is RtfNodeType.Control or RtfNodeType.Keyword)
                {
                    res.Append('\\');
                }
                else  //curNode.NodeType == RtfNodeType.Text
                {
                    if (prevNode is { NodeType: RtfNodeType.Keyword })
                    {
                        var code = char.ConvertToUtf32(curNode.NodeKey, 0);

                        if (code is >= 32 and < 128)
                            res.Append(' ');
                    }
                }

                AppendEncoded(res, curNode.NodeKey, enc);

                if (curNode.HasParameter)
                {
                    if (curNode.NodeType == RtfNodeType.Keyword)
                    {
                        res.Append(Convert.ToString(curNode.Parameter));
                    }
                    else if (curNode is { NodeType: RtfNodeType.Control, NodeKey: "\'" })
                        // If it is a special character like accented vowels
                    {						
                        res.Append(GetHex(curNode.Parameter));
                    }
                }

                break;
            }
        }

        // Child nodes are obtained
        var childrenLocal = curNode.ChildNodes;

        // If the node has children, the RTF code of the children is obtained.
        if (childrenLocal != null)
        {
            for (var i = 0; i < childrenLocal.Count; i++)
            {
                var node = childrenLocal[i];
                if (node == null) continue;

                if (i > 0)
                    res.Append(getRtfInm(node, childrenLocal[i - 1], enc));
                else
                    res.Append(getRtfInm(node, null, enc));
            }
        }

        if (curNode.NodeType == RtfNodeType.Group)
        {
            res.Append('}');
        }

        return res.ToString();
    }

    /// <summary>
    /// Concatenates two strings using the document encoding.
    /// </summary>
    /// <param name="res">Original text.</param>
    /// <param name="s">Text to append.</param>
    /// <param name="enc">Encoding of the document.</param>
    private void AppendEncoded(StringBuilder res, string s, Encoding? enc)
    {
        //Contributed by Jan Stuchlík
        enc ??= Tree?.GetEncoding() ?? Encoding.Default;
        
        for (var i = 0; i < s.Length; i++)
        {
            var code = char.ConvertToUtf32(s, i);

            if (code is >= 128 or < 32)
            {
                res.Append(@"\'");
                var bytes = enc.GetBytes([s[i]]);
                res.Append(GetHex(bytes[0]));
            }
            else
            {
                if ((s[i] == '{') || (s[i] == '}') || (s[i] == '\\'))
                {
                    res.Append('\\');
                }

                res.Append(s[i]);
            }
        }
    }

    /// <summary>
    /// Gets the hexadecimal code of an integer.
    /// </summary>
    /// <param name="code">Integer number.</param>
    /// <returns>Hexadecimal code of the integer passed as a parameter.</returns>
    private string GetHex(int code)
    {
        //Contributed by Jan Stuchlík

        var hex = Convert.ToString(code, 16);

        if (hex.Length == 1)
        {
            hex = "0" + hex;
        }

        return hex;
    }

    /// <summary>
    /// Updates the Root and Tree properties of a node (and its children) with those of the current node.
    /// </summary>
    /// <param name="node">Node to update.</param>
    private void updateNodeRoot(RtfTreeNode node)
    {
        // The root node of the document is assigned
        node.RootNode = RootNode;

        // The node's owner tree is assigned
        node.Tree = Tree;

        // If the updated node has children, they are updated as well.
        if (node.children == null) return;
        // Children of the current node are updated recursively
        foreach (RtfTreeNode nod in node.children)
        {
            updateNodeRoot(nod);
        }
    }

    /// <summary>
    /// Gets the text contained in the current node.
    /// </summary>
    /// <param name="raw">If this parameter is enabled, all text contained in the node will be extracted, regardless of whether it is part of the actual text of the document.</param>
    /// <param name="ignoreNchars">Ignore next N chars following \uN keyword</param>
    /// <returns>Text extracted from the node.</returns>
    private string GetText(bool raw, int ignoreNchars = 1)
    {
        var res = new StringBuilder("");

        switch (NodeType)
        {
            case RtfNodeType.Group:
            {
                var indkw = FirstChild is { NodeKey: "*" } ? 1 : 0;

                if (ChildNodes == null || ChildNodes[indkw] != null && (!raw &&
                                                             (ChildNodes[indkw]!.NodeKey.Equals("fonttbl") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("colortbl") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("stylesheet") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("generator") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("info") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("pict") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("object") ||
                                                              ChildNodes[indkw]!.NodeKey.Equals("fldinst")))) return res.ToString();
                if (ChildNodes != null)
                {
                    var uc = ignoreNchars;
                    foreach (RtfTreeNode node in ChildNodes)
                    {
                        res.Append(node.GetText(raw, uc));

                        if (node is { NodeType: RtfNodeType.Keyword, NodeKey: "uc" })
                            uc = node.Parameter;
                    }
                }

                break;
            }
            case RtfNodeType.Control when NodeKey == "'":
                res.Append(DecodeControlChar(Parameter, Tree?.GetEncoding() ?? Encoding.Default));
                break;
            case RtfNodeType.Control:
            {
                if (NodeKey == "~")  // non-breaking space
                    res.Append(' ');
                break;
            }
            case RtfNodeType.Text:
            {
                var newtext = NodeKey;

                // If the previous element was a Unicode character (\uN) we ignore the following N characters
                // according to the last tag \ucN
                if (PreviousNode is { NodeType: RtfNodeType.Keyword, NodeKey: "u" })
                {
                    newtext = newtext[ignoreNchars..];
                }

                res.Append(newtext);
                break;
            }
            case RtfNodeType.Keyword when NodeKey.Equals("par"):
                res.AppendLine("");
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("tab"):
                res.Append('\t');
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("line"):
                res.AppendLine("");
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("lquote"):
                res.Append('‘');
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("rquote"):
                res.Append('’');
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("ldblquote"):
                res.Append('“');
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("rdblquote"):
                res.Append('”');
                break;
            case RtfNodeType.Keyword when NodeKey.Equals("emdash"):
                res.Append('—');
                break;
            case RtfNodeType.Keyword:
            {
                if (NodeKey.Equals("u"))
                {
                    res.Append(char.ConvertFromUtf32(Parameter));
                }

                break;
            }
        }

        return res.ToString();
    }

    #endregion

    #region Properties

    /// <summary>
    /// Returns the root node of the document tree.
    /// </summary>
    /// <remarks>
    /// This is not the root node of the tree, but simply a dummy node of type ROOT from which the rest of the RTF tree branches off.
    /// It will therefore have only one child node of type GROUP, the real root of the tree.
    /// </remarks>
    public RtfTreeNode? RootNode { get; set; }

    /// <summary>
    /// Returns the parent node of the current node.
    /// </summary>
    public RtfTreeNode? ParentNode { get; set; }

    /// <summary>
    /// Returns the Rtf tree to which the node belongs.
    /// </summary>
    public RtfTree? Tree { get; set; }

    /// <summary>
    /// Returns the type of the current node.
    /// </summary>
    public RtfNodeType NodeType { get; set; }

    /// <summary>
    /// Returns the keyword, control symbol, or text of the current node.
    /// </summary>
    public string NodeKey { get; set; }

    /// <summary>
    /// Indicates whether the current node has a parameter assigned.
    /// </summary>
    public bool HasParameter { get; set; }

    /// <summary>
    /// Returns the parameter assigned to the current node.
    /// </summary>
    public int Parameter { get; set; }

    /// <summary>
    /// Returns the collection of child nodes of the current node.
    /// </summary>
    public RtfNodeCollection? ChildNodes
    {
        get => children;
        set
        {
            children = value;

            if (children == null) return;
            foreach (RtfTreeNode node in children)
            {
                node.ParentNode = this;

                // The Root and Tree properties of the new node and its possible children are updated
                updateNodeRoot(node);
            }
        }
    }

    /// <summary>
    /// Returns the first child node whose keyword is the one given as a parameter.
    /// </summary>
    /// <param name="keyword">Keyword searched.</param>
    /// <returns>First child node whose keyword is the one specified as a parameter. If it does not exist, null is returned.</returns>
    public RtfTreeNode? this[string keyword] => SelectSingleChildNode(keyword);

    /// <summary>
    /// Returns the n-th child of the current node.
    /// </summary>
    /// <param name="childIndex">Index of the child node to retrieve.</param>
    /// <returns>Nth child node of the current node. Returns null if it does not exist.</returns>
    public RtfTreeNode? this[int childIndex]
    {
        get
        { 
            RtfTreeNode? res = null;

            if (children != null && childIndex >= 0 && childIndex < children.Count)
                res = children[childIndex];

            return res;
        }
    }

    /// <summary>
    /// Returns the first child node of the current node.
    /// </summary>
    public RtfTreeNode? FirstChild
    {
        get
        {
            RtfTreeNode? res = null;

            if (children is { Count: > 0 })
                res = children[0];

            return res;
        }
    }

    /// <summary>
    /// Returns the last child node of the current node.
    /// </summary>
    public RtfTreeNode? LastChild
    {
        get
        {
            RtfTreeNode? res = null;

            return children is { Count: > 0 } ? children[^1] : res;
        }
    }

    /// <summary>
    /// Returns the next sibling node of the current node (Two nodes are siblings if they have the same parent node [ParentNode]).
    /// </summary>
    public RtfTreeNode? NextSibling
    {
        get
        {
            RtfTreeNode? res = null;

            if (ParentNode?.children == null) return res;
            
            var currentIndex = ParentNode.children.IndexOf(this);

            if (ParentNode.children.Count > currentIndex + 1)
                res = ParentNode.children[currentIndex + 1];

            return res;
        }
    }

    /// <summary>
    /// Returns the previous sibling node of the current node (Two nodes are siblings if they have the same parent node [ParentNode]).
    /// </summary>
    public RtfTreeNode? PreviousSibling
    {
        get
        {
            RtfTreeNode? res = null;

            if (ParentNode?.children == null) return res;
            var currentIndex = ParentNode.children.IndexOf(this);

            if (currentIndex > 0)
                res = ParentNode.children[currentIndex - 1];

            return res;
        }
    }

    /// <summary>
    /// Returns the next node in the tree.
    /// </summary>
    public RtfTreeNode? NextNode
    {
        get
        {
            RtfTreeNode? res = null;

            if (NodeType == RtfNodeType.Root)
            {
                res = FirstChild;
            }
            else if (ParentNode is { children: not null })
            {
                if (NodeType == RtfNodeType.Group && children is { Count: > 0 })
                {
                    res = FirstChild;
                }
                else
                {
                    if (Index < (ParentNode.children.Count - 1))
                    {
                        res = NextSibling;
                    }
                    else
                    {
                        res = ParentNode.NextSibling;
                    }
                }
            }

            return res;
        }
    }

    /// <summary>
    /// Returns the previous node in the tree.
    /// </summary>
    public RtfTreeNode? PreviousNode
    {
        get
        {
            RtfTreeNode? res = null;

            if (NodeType == RtfNodeType.Root)
            {
                res = null;
            }
            else if (ParentNode is { children: not null })
            {
                if (Index > 0)
                {
                    if (PreviousSibling is { NodeType: RtfNodeType.Group })
                    {
                        res = PreviousSibling.LastChild;
                    }
                    else
                    {
                        res = PreviousSibling;
                    }
                }
                else
                {
                    res = ParentNode;
                }
            }

            return res;
        }
    }

    /// <summary>
    /// Returns the RTF code of the current node and all its child nodes.
    /// </summary>
    public string Rtf => getRtf();

    /// <summary>
    /// Returns the index of the current node within the list of children of its parent node.
    /// </summary>
    public int Index
    {
        get
        {
            var res = -1;

            if(ParentNode is { children: not null }) 
                res = ParentNode.children.IndexOf(this);

            return res;
        }
    }

    /// <summary>
    /// Returns the text fragment of the document contained in the current node.
    /// </summary>
    public string Text => GetText(false);

    /// <summary>
    /// Returns all text contained in the current node.
    /// </summary>
    public string RawText => GetText(true);

    #endregion
}
