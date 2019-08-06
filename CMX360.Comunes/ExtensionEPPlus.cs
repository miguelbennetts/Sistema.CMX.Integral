using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Drawing;
using OfficeOpenXml.Drawing;
using System.Web;

using System.Windows.Forms;
using CMX360.Comunes.Clases;

namespace System
{


    public static class ExtensionEPPlus
    {
        public static string formatoDinero { get { return "$###,###,##0"; } }

        public static string formatoDineroSinSigno { get { return "###,###,##0"; } }
        public static string MonthName(this int month)
        {
            DateTimeFormatInfo dtinfo = new CultureInfo("es-ES", false).DateTimeFormat;
            return dtinfo.GetMonthName(month);
        }

        public static ExcelRange AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor)
        {
            ws.Cells[fila, columna].Value = valor;
            return ws.Cells[fila, columna];
        }

        public static ExcelRange ObtieneColumnas(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna)
        {

            ExcelRange rango = ws.Cells[fila, columna, aLafila, aLaColumna];
            return rango;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, string formato)
        {
            var cell = ws.Cells[fila, columna];
            cell.Value = valor;
            cell.Style.Numberformat.Format = formato;
        }
        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, string formato, bool fontBold, ExcelBorderStyle excelBorderStyle)
        {
            var cell = ws.Cells[fila, columna];
            cell.Value = valor;
            cell.Style.Numberformat.Format = formato;
            cell.Style.Font.Bold = fontBold;
            cell.Style.Border.Top.Style = excelBorderStyle;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, string formato, bool fontBold, ExcelBorderStyle excelBorderStyle, ExcelHorizontalAlignment horizontalAlignment)
        {
            var cell = ws.Cells[fila, columna];
            cell.Value = valor;
            cell.Style.Numberformat.Format = formato;
            cell.Style.Font.Bold = fontBold;
            cell.Style.Border.Top.Style = excelBorderStyle;
            cell.Style.HorizontalAlignment = horizontalAlignment;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, bool fontBold)
        {
            ws.Cells[fila, columna].Value = valor;
            ws.Cells[fila, columna].Style.Font.Bold = fontBold;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, bool fontBold, ExcelHorizontalAlignment horizontalAlignment)
        {
            ws.Cells[fila, columna].Value = valor;
            ws.Cells[fila, columna].Style.Font.Bold = fontBold;
            ws.Cells[fila, columna].Style.HorizontalAlignment = horizontalAlignment;
        }



        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, ExcelHorizontalAlignment horizontalAlignment)
        {
            ws.Cells[fila, columna].Value = valor;
            ws.Cells[fila, columna].Style.HorizontalAlignment = horizontalAlignment;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, object valor, ExcelHorizontalAlignment horizontalAlignment, bool fontBold)
        {
            ws.Cells[fila, columna].Value = valor;
            ws.Cells[fila, columna].Style.HorizontalAlignment = horizontalAlignment;
            ws.Cells[fila, columna].Style.Font.Bold = fontBold;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, string direccion, object valor)
        {
            ws.Cells[direccion].Value = valor;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, string direccion, object valor, bool fontBold)
        {
            ws.Cells[direccion].Value = valor;
            ws.Cells[direccion].Style.Font.Bold = fontBold;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, string direccion, object valor, ExcelHorizontalAlignment horizontalAlignment)
        {
            ws.Cells[direccion].Value = valor;
            ws.Cells[direccion].Style.HorizontalAlignment = horizontalAlignment;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, string direccion, object valor, ExcelHorizontalAlignment horizontalAlignment, bool fontBold)
        {
            ws.Cells[direccion].Value = valor;
            ws.Cells[direccion].Style.HorizontalAlignment = horizontalAlignment;
            ws.Cells[direccion].Style.Font.Bold = fontBold;
        }


        public static ExcelRange AgregarTexto(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna, object valor, float size)
        {
            var rango = ws.Cells[fila, columna, aLafila, aLaColumna];
            rango.Value = valor;
            rango.Style.Font.Size = size;
            return rango;
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna, object valor, float size, bool fontBold)
        {
            using (var rango = ws.Cells[fila, columna, aLafila, aLaColumna])
            {
                rango.Value = valor;
                rango.Style.Font.Bold = fontBold;
                rango.Style.Font.Size = size;
                rango.Merge = true;
            }
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna, object valor, float size, bool fontBold, ExcelHorizontalAlignment horizontalAlignment)
        {
            using (var rango = ws.Cells[fila, columna, aLafila, aLaColumna])
            {
                rango.Value = valor;
                rango.Style.Font.Bold = fontBold;
                rango.Style.Font.Size = size;
                rango.Merge = true;
                rango.Style.HorizontalAlignment = horizontalAlignment;
            }
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna, object valor, float size, bool fontBold, ExcelHorizontalAlignment horizontalAlignment, Color bgColor)
        {

            using (var rango = ws.Cells[fila, columna, aLafila, aLaColumna])
            {
                rango.Value = valor;
                rango.Style.Font.Bold = fontBold;
                rango.Style.Font.Size = size;
                rango.Merge = true;
                rango.Style.HorizontalAlignment = horizontalAlignment;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(bgColor);
            }
        }

        public static void AgregarTexto(this ExcelWorksheet ws, int fila, int columna, int aLafila, int aLaColumna, object valor, float size, bool fontBold, ExcelHorizontalAlignment horizontalAlignment, Color bgColor, bool ajustarTexto)
        {
            using (var rango = ws.Cells[fila, columna, aLafila, aLaColumna])
            {
                rango.Value = valor;
                rango.Style.Font.Bold = fontBold;
                rango.Style.Font.Size = size;
                rango.Merge = true;
                rango.Style.HorizontalAlignment = horizontalAlignment;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(bgColor);
                rango.Style.WrapText = ajustarTexto;
            }
        }


        public static void AgregarSuma(this ExcelWorksheet ws, int fila, int columna, int filaInicioSuma, ExcelBorderStyle excelBorderStyle)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = "Sum(" + ws.Cells[filaInicioSuma, columna].Address + ":" + ws.Cells[fila - 1, columna].Address + ")";
            cell.Style.Font.Bold = true;
            cell.Style.Border.Top.Style = excelBorderStyle;
            cell.Style.Numberformat.Format = "$###,###,##0.00";
        }

        public static ExcelRange AgregarSuma(this ExcelWorksheet ws, int fila, int columna, int filaInicioSuma, int filaFinSuma, ExcelBorderStyle excelBorderStyle)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = "Sum(" + ws.Cells[filaInicioSuma, columna].Address + ":" + ws.Cells[filaFinSuma - 1, columna].Address + ")";
            cell.Style.Font.Bold = true;
            cell.Style.Border.Top.Style = excelBorderStyle;
            cell.Style.Numberformat.Format = "$###,###,##0";
            return cell;
        }


        public static ExcelRange AgregarTotal(this ExcelRange cell, string Rango, ExcelBorderStyle excelBorderStyle, string formato)
        {

            cell.Formula = "Sum(" + Rango + ")";
            cell.Style.Font.Bold = true;
            cell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = cell.Style.Border.Bottom.Style = excelBorderStyle;
            cell.Style.Numberformat.Format = formato;
            return cell;
        }
        public static void AgregarSumaTotales(this ExcelWorksheet ws, int fila, int columna, string formula, ExcelBorderStyle excelBorderStyle)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = formula;
            cell.Style.Font.Bold = true;
            cell.Style.Border.Bottom.Style = excelBorderStyle;
            cell.Style.Numberformat.Format = "$###,###,##0";
        }

        public static void AgregarTotalPorcentaje(this ExcelWorksheet ws, int fila, int columna, object valor, ExcelBorderStyle borderStyleTop)
        {
            var cell = ws.Cells[fila, columna];
            cell.Value = valor;
            cell.Style.Font.Bold = true;
            cell.Style.Border.Top.Style = borderStyleTop;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }

        public static void AgregarTotalPorcentaje(this ExcelWorksheet ws, int fila, int columna, object valor, ExcelBorderStyle borderStyleTop, ExcelBorderStyle borderStyleBottom)
        {
            var cell = ws.Cells[fila, columna];
            cell.Value = valor;
            cell.Style.Font.Bold = true;
            cell.Style.Border.Top.Style = borderStyleTop;
            cell.Style.Border.Bottom.Style = borderStyleBottom;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }

        public static string ArmaFormulaSuma(this DatosSuma datos, ExcelWorksheet ws)
        {
            StringBuilder sbSuma = new StringBuilder();

            sbSuma.Append("Sum(");
            foreach (int fila in datos.Filas)
            {
                sbSuma.Append(ws.Cells[fila, datos.Columna].Address);
                sbSuma.Append("+");
            }

            sbSuma = sbSuma.Remove(sbSuma.Length - 1, 1);
            sbSuma.Append(")");

            return sbSuma.ToString();
        }

        public static ExcelWorksheet CrearHoja(this ExcelPackage p, string nombre, float size, string fuente)
        {
            p.Workbook.Worksheets.Add(nombre);
            ExcelWorksheet ws = p.Workbook.Worksheets[nombre];
            ws.Name = nombre;
            ws.Cells.Style.Font.Size = size;
            ws.Cells.Style.Font.Name = fuente;
            return ws;
        }

        public static ExcelWorksheet Fuente(this ExcelWorksheet ew, int fuente = 11, string nombreFuente = "Arial")
        {
            ew.Cells.Style.Font.Size = fuente;
            ew.Cells.Style.Font.Name = nombreFuente;
            return ew;
        }

        public static void AgregaPropiedadesWorkbook(this ExcelPackage p, string autor, string titulo)
        {
            p.Workbook.Properties.Author = autor;
            p.Workbook.Properties.Title = titulo;
        }

        public static void Titulo(this ExcelRange cell, string Titulo)
        {
            cell.Merge = true;
            cell.Style.Border.Top.Style = cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = cell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
            cell.Style.Font.Size = 8;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Value = Titulo;
            cell.Style.WrapText = true;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            var color = System.Drawing.ColorTranslator.FromHtml("#BFBFBF");
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        public static ExcelRange TextoRepProv(this ExcelRange cell, object texto, string formato)
        {
            cell.Value = texto;
            cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = ExcelBorderStyle.Double;
            cell.Style.Numberformat.Format = formato;

            return cell;
        }

        public static ExcelRange FontSize(this ExcelRange cell, int tamanio)
        {
            cell.Style.Font.Size = tamanio;
            return cell;
        }

        public static ExcelRange TextoRepProv(this ExcelRange cell, object texto)
        {
            return cell.TextoRepProv(texto, string.Empty);
        }

        public static ExcelRange Combinar(this ExcelRange cell)
        {
            cell.Merge = true;
            return cell;
        }

        public static ExcelRange PorcentajeProvRep(this ExcelRange cell, string celda, string celdaTotal)
        {
            cell.Formula = string.Format("IFERROR({0}/{1},0)", celda, celdaTotal);
            cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = ExcelBorderStyle.Double;
            cell.Style.Numberformat.Format = "0.0%";

            return cell;
        }
        public static ExcelRange PorcentajeProvRep(this ExcelRange cell, string celda, string celdaTotal, string error)
        {
            cell.Formula = string.Format("IFERROR({0}/{1},{2})", celda, celdaTotal, error);
            cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = ExcelBorderStyle.Double;
            cell.Style.Numberformat.Format = "0.0%";
            return cell;
        }

        public static ExcelRange BordeCompleto(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Top.Style = cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = cell.Style.Border.Bottom.Style = tipoBorde;
            return cell;
        }
        public static ExcelRange BordeArriba(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Top.Style = tipoBorde;
            return cell;
        }
        public static ExcelRange BordeAbajo(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Bottom.Style = tipoBorde;
            return cell;
        }
        public static ExcelRange Fondo(this ExcelRange cell, string strColor = "#BFBFBF")
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            var color = System.Drawing.ColorTranslator.FromHtml(strColor);
            cell.Style.Fill.BackgroundColor.SetColor(color);
            return cell;
        }

        public static ExcelRange Negrita(this ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            return cell;
        }
        public static ExcelRange Alinea(this ExcelRange cell, ExcelHorizontalAlignment Alinea = ExcelHorizontalAlignment.Center)
        {
            cell.Style.HorizontalAlignment = Alinea;
            return cell;
        }
        public static ExcelRange Resta(this ExcelRange cell, string celda, string celdaTotal)
        {
            cell.Formula = string.Format("IFERROR({0}-{1},0)", celda, celdaTotal);
            cell.Style.Border.Left.Style = cell.Style.Border.Right.Style = ExcelBorderStyle.Double;
            cell.Style.Numberformat.Format = formatoDinero;
            return cell;
        }



        public static ExcelRange FormatoNumero(this ExcelRange cell, string formato)
        {
            cell.Style.Numberformat.Format = formato;
            return cell;
        }

        public static ExcelRange HorizontalAlign(this ExcelRange cell, ExcelHorizontalAlignment align = ExcelHorizontalAlignment.Right)
        {
            cell.Style.HorizontalAlignment = align;
            return cell;
        }

        public static ExcelRange VerticalAlign(this ExcelRange cell, ExcelVerticalAlignment align = ExcelVerticalAlignment.Top)
        {
            cell.Style.VerticalAlignment = align;
            return cell;
        }

        public static ExcelRange FontBold(this ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            return cell;
        }

        public static ExcelRange borderTop(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Top.Style = tipoBorde;
            return cell;
        }

        public static ExcelRange borderBottom(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Bottom.Style = tipoBorde;
            return cell;
        }

        public static ExcelRange borderLeft(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Left.Style = tipoBorde;
            return cell;
        }
        public static ExcelRange borderRight(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Right.Style = tipoBorde;
            return cell;
        }

        public static ExcelRange borderLeftRight(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Left.Style = tipoBorde;
            cell.Style.Border.Right.Style = tipoBorde;
            return cell;
        }

        public static ExcelRange borderTopBottom(this ExcelRange cell, ExcelBorderStyle tipoBorde = ExcelBorderStyle.Double)
        {
            cell.Style.Border.Top.Style = tipoBorde;
            cell.Style.Border.Bottom.Style = tipoBorde;
            return cell;
        }

        public static ExcelRange AjustarTexto(this ExcelRange cell)
        {
            cell.Style.WrapText = true;
            return cell;
        }

        public static ExcelRange ColorFuente(this ExcelRange cell, string strColor = "#000000")
        {
            var color = System.Drawing.ColorTranslator.FromHtml(strColor);
            //cell.Style.Fill.PatternColor.SetColor(color);
            cell.Style.Font.Color.SetColor(color);
            return cell;
        }


        public static ExcelRange AgregarSuma(this ExcelWorksheet ws, int fila, int columna, int filaInicioSuma, int filaFinSuma)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = "Sum(" + ws.Cells[filaInicioSuma, columna].Address + ":" + ws.Cells[filaFinSuma - 1, columna].Address + ")";
            cell.Style.Numberformat.Format = "$###,###,##0";
            return cell;
        }

        public static ExcelRange AgregarResta(this ExcelWorksheet ws, int fila, int columna, int filaMinuendo, int filaSustraendo)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = "=" + ws.Cells[filaMinuendo, columna].Address + "-" + ws.Cells[filaSustraendo, columna].Address;
            return cell;
        }

        public static ExcelRange AgregarFormula(this ExcelWorksheet ws, int fila, int columna, string formula)
        {
            var cell = ws.Cells[fila, columna];
            cell.Formula = formula;
            return cell;
        }

        public static ExcelRange Dividir(this ExcelRange cell, string celda, string celdaTotal, int digitos = 0)
        {
            cell.Formula = string.Format("=ROUND(IFERROR({0}/{1},0),{2})", celda, celdaTotal, digitos);
            return cell;
        }


        public static ExcelPicture AgregaImagen(this ExcelWorksheet ws, int fila, int columna, int size = 100)
        {
            string ruta = "";
            if (HttpContext.Current != null)
                ruta = HttpContext.Current.Server.MapPath("~/");
            else
                ruta = Application.StartupPath + @"\";
            var archivo = @"images\FarmaciaParisRojo.jpg";
            var rutaArchivo = string.Format("{0}{1}", ruta, archivo);

            Bitmap b = new Bitmap(rutaArchivo);
            ExcelPicture imagen = null;
            if (b != null)
            {
                imagen = ws.Drawings.AddPicture("pic" + fila.ToString() + columna.ToString(), b);
                imagen.From.Column = columna;
                imagen.From.Row = fila;
                imagen.SetSize(size, size);
            }
            return imagen;
        }

        public static ExcelPicture EspacioImagen(this ExcelPicture imagen, int pixel)
        {
            imagen.From.ColumnOff = pixel * 9525;
            imagen.From.RowOff = pixel * 9525;
            return imagen;
        }

        public static ExcelPicture SizeImagen(this ExcelPicture imagen, int w, int h)
        {
            imagen.SetSize(w, h);
            return imagen;
        }

        public static ExcelRange AgregaFormula(this ExcelRange cell, string formula)
        {
            cell.Formula = formula;
            return cell;
        }

    }


}
