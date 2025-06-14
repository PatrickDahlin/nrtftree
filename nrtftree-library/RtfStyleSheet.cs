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
 * Class:		RtfStyleSheet
 * Description:	Representa una hoja de estilo de un documento RTF.
 * ******************************************************************************/

using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree.Util;


/// <summary>
/// Representa una hoja de estilo de un documento RTF
/// </summary>
public class RtfStyleSheet
{

    /// <summary>
    /// Índice de la hoja de estilo
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Nombre de la hoja de estilo.
    /// </summary>
    public string Name { get; set; } = "";

    /// <summary>
    /// Tipo de la hoja de estilo.
    /// </summary>
    public RtfStyleSheetType Type { get; set; } = RtfStyleSheetType.Paragraph;

    /// <summary>
    /// En las hojas de estilo de tipo Caracter indica si el estilo indicado se va a sumar al estilo de párrafo actual, en vez de sobreescribir completamente el estilo actual.
    /// </summary>
    public bool Additive { get; set; }

    /// <summary>
    /// Indica el estilo en el que está basado el estilo actual.
    /// </summary>
    public int BasedOn { get; set; } = -1;

    /// <summary>
    /// Indica el estilo que se usará en el siguiente párrafo.
    /// </summary>
    public int Next { get; set; } = -1;

    /// <summary>
    /// Indican si los estilos se actualizarán automáticamente.
    /// </summary>
    public bool AutoUpdate { get; set; }

    /// <summary>
    /// Establece que el estilo actual no aparecerá en las listas desplegables de estilos.
    /// </summary>
    public bool Hidden { get; set; }

    /// <summary>
    /// Indica el estilo al que está enlazado el actual.
    /// </summary>
    public int Link { get; set; } = -1;

    /// <summary>
    /// Indica que el estilo actual está bloqueado.
    /// </summary>
    public bool Locked { get; set; }

    /// <summary>
    /// Indica que el estilo actual es el estilo de e-mail personal.
    /// </summary>
    public bool Personal { get; set; }

    /// <summary>
    /// Indica que el estilo actual es el estilo de composición de e-mail.
    /// </summary>
    public bool Compose { get; set; }

    /// <summary>
    /// Indica que el estilo actual es el estilo de respuesta de e-mail.
    /// </summary>
    public bool Reply { get; set; }

    /// <summary>
    /// Tied to the rsid table, N is the rsid of the author who implemented the style.
    /// </summary>
    public int Styrsid { get; set; } = -1;

    /// <summary>
    /// Indica que el estilo no aparecerá en los menús desplegables.
    /// </summary>
    public bool SemiHidden { get; set; } = false;

    /// <summary>
    /// Indica la tecla rápida para establecer el estilo actual.
    /// </summary>
    public RtfNodeCollection? KeyCode { get; set; }

    /// <summary>
    /// Opciones de formato contenidas en el estilo actual.
    /// </summary>
    public RtfNodeCollection? Formatting { get; set; }

    /// <summary>
    /// Constructor.
    /// </summary>
    public RtfStyleSheet()
    {
        KeyCode = null;
        Formatting = new RtfNodeCollection();
    }

}
