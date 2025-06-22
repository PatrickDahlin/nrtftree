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
 * Class:		Rtf2Html
 * Description:	Traducción de documentos RTF a formato HTML.
 * Notes:       Contribución de Francisco Javier Marín (http://www.xuletas.es/).
 ********************************************************************************/

using System.Text;
using System.Text.RegularExpressions;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;
using SkiaSharp;

namespace Net.Sgoliver.NRtfTree
{
    namespace Demo
    {
        /// <summary>
        /// Conversor de documentos RTF a formato HTML.
        /// </summary>
        public partial class Rtf2Html
        {
            #region Atributos Privados

            /// <summary>
            /// StringBuilder empleado para escribir el código HTML resultado
            /// </summary>
            private StringBuilder _builder;

            /// <summary>
            /// Formato de texto que se está leyendo
            /// </summary>
            private Format _currentFormat;

            /// <summary>
            /// Formato de texto ya escrito en el código HTML
            /// </summary>
            private Format _htmlFormat;

            /// <summary>
            /// Tabla de colores RTF
            /// </summary>
            private RtfColorTable _colorTable;

            /// <summary>
            /// Tabla de fuentes RTF
            /// </summary>
            private RtfFontTable _fontTable;

            //Propiedades

            #endregion

            #region Constructores

            public Rtf2Html()
            {
                AutoParagraph = true;
                IgnoreFontNames = false;
                DefaultFontSize = 10;
                EscapeHtmlEntities = true;
                DefaultFontName = "Times New Roman";

                ImageFormat = ImageFormat.Jpeg;
                EmbedImages = true;
                ImagePath = "";
            }

            #endregion

            #region Propiedades

            /// <summary>
            /// Obtiene o establece un valor que indica si se crearán parrafos automáticamente
            /// </summary>
            public bool AutoParagraph { get; set; }

            /// <summary>
            /// Obtiene o establece un valor que indica si se ignorarán los nombres de las fuentes.
            /// Establecer este valor como false generará un archivo HTML sin propiedad CSS font-face
            /// </summary>
            public bool IgnoreFontNames { get; set; }

            /// <summary>
            /// Obtiene o establece un valor que indica si se reemplazarán las entidades HTML
            /// del texto
            /// </summary>
            public bool EscapeHtmlEntities { get; set; }

            /// <summary>
            /// Obtiene o establece un valor que indica el nombre de la fuente por defecto.
            /// Los grupos de textos que usen esta fuente no incluirán la propiedad font-face
            /// </summary>
            public string DefaultFontName { get; set; }

            /// <summary>
            /// Obtiene o establece un valor que indica el tamaño de fuente que se ignorará.
            /// Los grupos de texto que tengan ese tamaño no incluirán la propiedad CSS font-size
            /// </summary>
            public int DefaultFontSize { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether to embed base64 images.
            /// within the HTML code (true), or they will be saved to a file (false)
            /// </summary>
            /// <see cref="http://www.sweeting.org/mark/blog/2005/07/12/base64-encoded-images-embedded-in-html"/>
            public bool EmbedImages { get; set; }

            /// <summary>
            /// Obtiene o establece la ruta donde se guardarán las imágenes contenidas en el
            /// código RTF original. Sólo se usarará si el valor de IncrustImages es false
            /// </summary>
            public string ImagePath { get; set; }

            /// <summary>
            /// Obtiene o establece un valor que indica el formato en el que se guardarán las imágenes
            /// incrustadas en el código RTF convertido
            /// </summary>
            public ImageFormat ImageFormat { get; set; }

            #endregion

            #region Métodos Públicos

            /// <summary>
            /// Convierte una cadena de código RTF a formato HTML
            /// </summary>
            public static string ConvertCode(string rtf)
            {
                var rtfToHtml = new Rtf2Html();
                return rtfToHtml.Convert(rtf);
            }

            /// <summary>
            /// Convierte una cadena de código RTF a formato HTML
            /// </summary>
            public string Convert(string rtf)
            {
                //Generar arbol DOM
                var rtfTree = new RtfTree();
                rtfTree.LoadRtfText(rtf);

                //Inicializar variables empleadas
                _builder = new StringBuilder();
                _htmlFormat = new Format();
                _currentFormat = new Format();
                _fontTable = rtfTree.GetFontTable();
                _colorTable = rtfTree.GetColorTable();

                //Buscar el inicio del contenido visible del documento
                int inicio;
                for (inicio = 0; inicio < rtfTree.RootNode.FirstChild.ChildNodes.Count; inicio++)
                {
                    if (rtfTree.RootNode.FirstChild.ChildNodes[inicio].NodeKey == "pard")
                        break;
                }

                //Procesar todos los nodos visibles
                ProcessChildNodes(rtfTree.RootNode.FirstChild.ChildNodes, inicio);

                //Cerrar etiquetas pendientes
                _currentFormat.Reset();
                WriteText(string.Empty);

                //Arreglar HTML

                //Arreglar listas
                var repairList = MyRegex1();

                foreach (Match match in repairList.Matches(_builder.ToString()))
                {
                    _builder.Replace(match.Value, string.Format("<li style=\"{0}\">{1}</li>", match.Groups[1].Value, match.Groups[2].Value));
                }

                var repairUl = new Regex("(?<!</li>)<li", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                foreach (Match match in repairUl.Matches(_builder.ToString()))
                {
                    _builder.Insert(match.Index, "<ul>");
                }

                repairUl = MyRegex();

                foreach (Match match in repairUl.Matches(_builder.ToString()))
                {
                    _builder.Insert(match.Index + match.Length, "</ul>");
                }

                //Generar párrafos (cada 2 <br /><br /> se cambiará por un <p>)
                if (AutoParagraph)
                {
                    var partes = _builder.ToString().Split(["<br /><br />"], StringSplitOptions.RemoveEmptyEntries);
                    _builder = new StringBuilder(_builder.Length + 7 * partes.Length);

                    foreach (var parte in partes)
                    {
                        _builder.Append("<p>");
                        _builder.Append(parte);
                        _builder.Append("</p>");
                    }
                }

                return EscapeHtmlEntities ? HtmlEntities.Encode(_builder.ToString()) : _builder.ToString();
            }

            #endregion

            #region Métodos Privados

            /// <summary>
            /// Función encargada de procesar los nodos hijo de un nodo padre RTF un nodo del documento RTF
            /// y generar las etiquetas HTML necesarias
            /// </summary>
            private void ProcessChildNodes(RtfNodeCollection nodos, int inicio)
            {
                for (var i = inicio; i < nodos.Count; i++)
                {
                    var node = nodos[i];

                    switch (node?.NodeType)
                    {
                        case RtfNodeType.Control:

                            if (node.NodeKey == "'") //Símbolos especiales, como tildes y "ñ"
                            {
                                WriteText(Encoding.Default.GetString(new[] { (byte)node.Parameter }));
                            }
                            break;

                        case RtfNodeType.Keyword:

                            switch (node.NodeKey)
                            {
                                case "pard": //Reinicio de formato
                                    _currentFormat.Reset();
                                    break;

                                case "f": //Tipo de fuente                                
                                    if (node.Parameter < _fontTable.Count)
                                        _currentFormat.FontName = _fontTable[node.Parameter];
                                    break;

                                case "cf": //Color de fuente
                                    if (node.Parameter < _colorTable.Count)
                                        _currentFormat.ForeColor = _colorTable[node.Parameter];
                                    break;

                                case "highlight": //Color de fondo
                                    if (node.Parameter < _colorTable.Count)
                                        _currentFormat.BackColor = _colorTable[node.Parameter];
                                    break;

                                case "fs": //Tamaño de fuente
                                    _currentFormat.FontSize = node.Parameter;
                                    break;

                                case "b": //Negrita
                                    _currentFormat.Bold = !node.HasParameter || node.Parameter == 1;
                                    break;

                                case "i": //Cursiva
                                    _currentFormat.Italic = !node.HasParameter || node.Parameter == 1;
                                    break;

                                case "ul": //Subrayado ON
                                    _currentFormat.Underline = true;
                                    break;

                                case "ulnone": //Subrayado OFF
                                    _currentFormat.Underline = false;
                                    break;

                                case "super": //Superscript
                                    _currentFormat.Superscript = true;
                                    _currentFormat.Subscript = false;
                                    break;

                                case "sub": //Subindice
                                    _currentFormat.Subscript = true;
                                    _currentFormat.Superscript = false;
                                    break;

                                case "nosupersub":
                                    _currentFormat.Superscript = _currentFormat.Subscript = false;
                                    break;

                                case "qc": //Alineacion centrada
                                    _currentFormat.Alignment = HorizontalAlignment.Center;
                                    break;

                                case "qr": //Alineacion derecha
                                    _currentFormat.Alignment = HorizontalAlignment.Right;
                                    break;

                                case "li": //tabulacion
                                    _currentFormat.Margin = node.Parameter;
                                    break;

                                case "line":
                                case "par": //Nueva línea
                                    _builder.Append("<br />");
                                    break;
                            }

                            break;

                        case RtfNodeType.Group:

                            //Procesar nodos hijo, si los tiene
                            if (node.HasChildNodes())
                            {
                                if (node["pict"] != null) //El grupo es una imagen
                                {
                                    var imageNode = new ImageNode(node);
                                    WriteImage(imageNode);
                                }
                                else
                                {
                                    ProcessChildNodes(node.ChildNodes, 0);
                                }
                            }
                            break;

                        case RtfNodeType.Text:

                            WriteText(node.NodeKey);
                            break;

                        default:

                            throw new ArgumentOutOfRangeException();
                    }
                }
            }

            private void WriteText(string text)
            {
                if (_builder.Length > 0)
                {
                    if (_currentFormat.Bold == false && _htmlFormat.Bold)
                    {
                        _builder.Append("</strong>");
                        _htmlFormat.Bold = false;
                    }
                    if (_currentFormat.Italic == false && _htmlFormat.Italic)
                    {
                        _builder.Append("</em>");
                        _htmlFormat.Italic = false;
                    }
                    if (_currentFormat.Underline == false && _htmlFormat.Underline)
                    {
                        _builder.Append("</u>");
                        _htmlFormat.Underline = false;
                    }
                    if (_currentFormat.Subscript == false && _htmlFormat.Subscript)
                    {
                        _builder.Append("</sub>");
                        _htmlFormat.Subscript = false;
                    }
                    if (_currentFormat.Superscript == false && _htmlFormat.Superscript)
                    {
                        _builder.Append("</sup>");
                        _htmlFormat.Superscript = false;
                    }
                    if (_currentFormat.CompareFontFormat(_htmlFormat) == false) //El formato de fuente ha cambiado
                    {
                        _builder.Append("</span>");

                        //Reiniciar propiedades
                        _htmlFormat.Reset();
                    }
                }

                //Abrir etiquetas necesarias para representar el formato actual
                if (_currentFormat.CompareFontFormat(_htmlFormat) == false) //El formato de fuente ha cambiado
                {
                    var format = string.Empty;

                    if (!IgnoreFontNames && !string.IsNullOrEmpty(_currentFormat.FontName) &&
                        string.Compare(_currentFormat.FontName, DefaultFontName, StringComparison.OrdinalIgnoreCase) != 0)
                        format += $"font-family:{_currentFormat.FontName};";
                    if (_currentFormat.FontSize > 0 && _currentFormat.FontSize / 2 != DefaultFontSize)
                        format += $"font-size:{_currentFormat.FontSize / 2}pt;";
                    if (_currentFormat.Margin != _htmlFormat.Margin)
                        format += $"margin-left:{_currentFormat.Margin / 15}px;";
                    if (_currentFormat.Alignment != _htmlFormat.Alignment)
                        format += $"text-align:{_currentFormat.Alignment.ToString().ToLower()};";
                    if (CompareColor(_currentFormat.ForeColor, _htmlFormat.ForeColor) == false)
                        format += $"color:{_currentFormat.ForeColor};";
                    if (CompareColor(_currentFormat.BackColor, _htmlFormat.BackColor) == false)
                        format += $"background-color:{_currentFormat.BackColor};";
                    
                    _htmlFormat.FontName = _currentFormat.FontName;
                    _htmlFormat.FontSize = _currentFormat.FontSize;
                    _htmlFormat.ForeColor = _currentFormat.ForeColor;
                    _htmlFormat.BackColor = _currentFormat.BackColor;
                    _htmlFormat.Margin = _currentFormat.Margin;
                    _htmlFormat.Alignment = _currentFormat.Alignment;

                    if (!string.IsNullOrEmpty(format))
                        _builder.Append($"<span style=\"{format}\">");
                }
                if (_currentFormat.Superscript && _htmlFormat.Superscript == false)
                {
                    _builder.Append("<sup>");
                    _htmlFormat.Superscript = true;
                }
                if (_currentFormat.Subscript && _htmlFormat.Subscript == false)
                {
                    _builder.Append("<sub>");
                    _htmlFormat.Subscript = true;
                }
                if (_currentFormat.Underline && _htmlFormat.Underline == false)
                {
                    _builder.Append("<u>");
                    _htmlFormat.Underline = true;
                }
                if (_currentFormat.Italic && _htmlFormat.Italic == false)
                {
                    _builder.Append("<em>");
                    _htmlFormat.Italic = true;
                }
                if (_currentFormat.Bold && _htmlFormat.Bold == false)
                {
                    _builder.Append("<strong>");
                    _htmlFormat.Bold = true;
                }

                _builder.Append(text.Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;"));
            }

            /// <summary>
            /// Función encargada de añadir una imagen al código HTML resultado
            /// </summary>
            /// <param name="imageNode"></param>
            private void WriteImage(ImageNode imageNode)
            {
                var b = imageNode.Bitmap;

                var imageSize = imageNode is { DesiredHeight: > 0, DesiredWidth: > 0 } 
                    ? new SKSize(imageNode.DesiredWidth / 15, imageNode.DesiredHeight / 15) 
                    : b.Info.Size;

                // Reduce image if final size is smaller than current size
                if (b.Info.Size.Height > imageSize.Height || b.Info.Size.Width > imageSize.Width)
                {
                    // Original was bicubic, but skia doesnt seem to support bicubic out of the box
                    var bmpResized = b.Resize(imageSize.ToSizeI(), new SKSamplingOptions(SKFilterMode.Nearest, SKMipmapMode.Nearest));
                    b.Dispose();
                    b = bmpResized;
                }

                string imageSrc;
                if (EmbedImages)
                {
                    SKEncodedImageFormat format;
                    string mimeType;
                    switch (ImageFormat)
                    {
                        case ImageFormat.Png: format = SKEncodedImageFormat.Png; mimeType = "image/png"; break;
                        case ImageFormat.Jpeg: format = SKEncodedImageFormat.Jpeg; mimeType = "image/jpeg"; break;
                        case ImageFormat.Bmp: format = SKEncodedImageFormat.Bmp; mimeType = "image/bmp"; break;
                        default:
                        case ImageFormat.Unknown: format = SKEncodedImageFormat.Png; mimeType = "image/png"; break;
                    }
                    var data = b.Encode(format, 90);

                    imageSrc = $"data:{mimeType};base64,{System.Convert.ToBase64String(data.Span)}";
                }
                else
                {
                    SKEncodedImageFormat format;
                    string ext;
                    switch (ImageFormat)
                    {
                        case ImageFormat.Png: format = SKEncodedImageFormat.Png;
                            ext = "png"; break;
                        case ImageFormat.Jpeg: format = SKEncodedImageFormat.Jpeg;
                            ext = "jpeg"; break;
                        case ImageFormat.Bmp: format = SKEncodedImageFormat.Bmp;
                            ext = "bmp"; break;
                        default:
                        case ImageFormat.Unknown: format = SKEncodedImageFormat.Png;
                            ext = "png"; break;
                    }
                    string path;
                    var index = 1;
                    do
                    {
                        path = Path.Combine(ImagePath, $"Imagen{index}.{ext}");
                        index++;
                    } while (File.Exists(path));

                    var data = b.Encode(format, 90);
                    File.WriteAllBytes(path, data.ToArray());

                    imageSrc = path;
                }

                _builder.Append($"<img src=\"{imageSrc}\" style=\"width:{imageSize.Width}px;height:{imageSize.Height}px;\" />");

                b.Dispose();
            }

            /// <summary>
            /// Compara dos colores sin tener en cuenta el canal alfa
            /// </summary>
            private static bool CompareColor(SKColor a, SKColor b)
            {
                return a.Red == b.Red && a.Green == b.Green && a.Blue == b.Blue;
            }

            #endregion

            #region Nested type: Format

            /// <summary>
            /// Representa el formato que puede tomar un conjunto de texto
            /// </summary>
            private class Format
            {
                public bool Italic;
                public bool Bold;
                public bool Subscript;
                public bool Underline;
                public bool Superscript;

                public string FontName;
                public int FontSize;
                public SKColor ForeColor;
                public SKColor BackColor;
                public int Margin;
                public HorizontalAlignment Alignment;

                public Format()
                {
                    Reset();
                }

                /// <summary>
                /// Compara las propiedades FontName, FontSize, Margin, ForeColor, BackColor y Alignment con otro
                /// objeto de esta clase
                /// </summary>
                public bool CompareFontFormat(Format format)
                {
                    return string.Compare(FontName, format.FontName, StringComparison.OrdinalIgnoreCase) == 0 &&
                           FontSize == format.FontSize &&
                           ForeColor == format.ForeColor &&
                           BackColor == format.BackColor &&
                           Margin == format.Margin &&
                           Alignment == format.Alignment;
                }

                public void Reset()
                {
                    FontName = string.Empty;
                    FontSize = 0;
                    ForeColor = SKColors.Black;
                    BackColor = SKColors.White;
                    Margin = 0;
                    Alignment = HorizontalAlignment.Left;
                }
            }

            private enum HorizontalAlignment
            {
                Left,
                Right,
                Center
            }

            [GeneratedRegex("/li>(?!<li)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
            private static partial Regex MyRegex();
            [GeneratedRegex("<span [^>]*>·</span><span style=\"([^\"]*)\">(.*?)<br\\s+/></span>", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant)]
            private static partial Regex MyRegex1();

            #endregion
        }
    }
}