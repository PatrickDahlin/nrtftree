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
 * Class:		RtfNodeCollection
 * Description:	Collection of nodes in an RTF document.
 * ******************************************************************************/

using System.Collections;

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Collection of nodes in an RTF document.
/// </summary>
public class RtfNodeCollection : CollectionBase
{
    #region Public Methods

    /// <summary>
    /// Adds a new node to the current collection.
    /// </summary>
    /// <param name="node">New node to add.</param>
    /// <returns>Position of the node in the collection.</returns>
    public int Add(RtfTreeNode node)
    {
        InnerList.Add(node);

        return (InnerList.Count - 1);
    }

    /// <summary>
    /// Insert new node at a specific position in the collection
    /// </summary>
    /// <param name="index">Position at which to insert new node.</param>
    /// <param name="node">Node to insert.</param>
    public void Insert(int index, RtfTreeNode node)
    {
        InnerList.Insert(index, node);
    }

    /// <summary>
    /// Indexer to get node at specific index
    /// </summary>
    public RtfTreeNode? this[int index]
    {
        get => (RtfTreeNode?)InnerList[index];
        set => InnerList[index] = value;
    }

    /// <summary>
    /// Returns the index of the node given in this collection
    /// </summary>
    /// <param name="node">Node to look for in the collection.</param>
    /// <returns>Index of the node, -1 if the node was not found in the collection.</returns>
    public int IndexOf(RtfTreeNode node)
    {
        return InnerList.IndexOf(node);
    }

    /// <summary>
    /// Returns the index of the node passed as a parameter within the list of nodes in the collection.
    /// </summary>
    /// <param name="node">Node to search for in the collection.</param>
    /// <param name="startIndex">Position within the collection from which the search will be performed.</param>
    /// <returns>Index of the node being searched for. Returns -1 if the node is not found in the collection.</returns>
    public int IndexOf(RtfTreeNode node, int startIndex)
    {
        return InnerList.IndexOf(node, startIndex);
    }

    /// <summary>
    /// Returns the index of the first node in the collection whose key is the one passed as a parameter.
    /// </summary>
    /// <param name="key">Key to search in the collection.</param>
    /// <returns>Index of the node being searched for. Returns -1 if the node is not found in the collection.</returns>
    public int IndexOf(string key)
    {
        var intFoundAt = -1;

        if (InnerList.Count <= 0) return intFoundAt;
        
        for (var intIndex = 0; intIndex < InnerList.Count; intIndex++)
        {
            if (((RtfTreeNode?)InnerList[intIndex])?.NodeKey != key) continue;
            
            intFoundAt = intIndex;
            break;
        }

        return intFoundAt;
    }

    /// <summary>
    /// Returns the index of the first node in the collection whose key is the one passed as a parameter.
    /// </summary>
    /// <param name="key">Key to search in the collection.</param>
    /// <param name="startIndex">Position within the collection from which the search will be performed.</param>
    /// <returns>Index of the node being searched for. Returns -1 if the node is not found in the collection.</returns>
    public int IndexOf(string key, int startIndex)
    {
        var intFoundAt = -1;

        if (InnerList.Count <= 0) return intFoundAt;
        
        for (var intIndex = startIndex; intIndex < InnerList.Count; intIndex++)
        {
            if (((RtfTreeNode?)InnerList[intIndex])?.NodeKey != key) continue;
            
            intFoundAt = intIndex;
            break;
        }

        return intFoundAt;
    }

    /// <summary>
    /// Adds a new list of nodes to the end of the collection.
    /// </summary>
    /// <param name="collection">New list of nodes to add to the current collection.</param>
    public void AddRange(RtfNodeCollection collection)
    {
        InnerList.AddRange(collection);
    }

    /// <summary>
    /// Removes a set of adjacent nodes from the collection.
    /// </summary>
    /// <param name="index">Index of the first node in the set to be removed.</param>
    /// <param name="count">Number of nodes to remove.</param>
    public void RemoveRange(int index, int count)
    {
        InnerList.RemoveRange(index, count);
    }

    #endregion
}
