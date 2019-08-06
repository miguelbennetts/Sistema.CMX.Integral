
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Dynamic;
using System.Security.Cryptography;

namespace System
{
    public static class Extensions
    {

        public static string toHTML_Table<T>(this T[] _lst)
        {

            StringBuilder builder = new StringBuilder();

            builder.Append("<center><table border='1px' cellpadding='5' cellspacing='0' ");
            builder.Append("style='border: solid 1px Silver; font-size: x-small;'>");
            builder.Append("<tr align='left' valign='top'>");
            foreach (var it in _lst.First().GetType().GetProperties())
            {
                builder.Append("<td align='left' valign='top'><b>");
                builder.Append(it.Name);
                builder.Append("</b></td>");
            }
            builder.Append("</tr>");
            foreach (var r in _lst)
            {
                builder.Append("<tr align='left' valign='top'>");
                foreach (var c in r.GetType().GetProperties())
                {
                    builder.Append("<td align='left' valign='top'>");
                    builder.Append(r.ValorPropiedad(c.Name));
                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</table></center>");

            return builder.ToString();
        }

        public static void ToExcel<T>(this T[] _lst, ref ExcelWorksheet _ws)
        {
            int fila = 1, columna = 1, AnchoCol = 0;

            foreach (var it in _lst.First().GetType().GetProperties())
            {
                _ws.Cells[fila, columna].Value = it.Name;

                _ws.Cells[fila, columna].borderTop(ExcelBorderStyle.Thin);
                _ws.Cells[fila, columna].borderBottom(ExcelBorderStyle.Thin);
                _ws.Cells[fila, columna].borderRight(ExcelBorderStyle.Thin);
                _ws.Cells[fila, columna].borderLeft(ExcelBorderStyle.Thin);

                _ws.Cells[fila, columna].Fondo("#B40404");
                _ws.Cells[fila, columna].ColorFuente("#FFFFFF");
                _ws.Cells[fila, columna].FontBold();

                _ws.Cells[fila, columna].AutoFitColumns();

                columna++;
            }

            foreach (var r in _lst)
            {
                columna = 1;
                fila++;
                foreach (var c in r.GetType().GetProperties())
                {
                    AnchoCol = 0;

                    if (r.ValorPropiedad(c.Name) is DateTime)
                    {
                        _ws.Cells[fila, columna].FormatoNumero("MM/DD/YYYY");
                    }
                    else if (r.ValorPropiedad(c.Name) is Int32 || r.ValorPropiedad(c.Name) is Int64)
                    {
                        _ws.Cells[fila, columna].FormatoNumero("0");
                    }
                    else if (r.ValorPropiedad(c.Name) is Double || r.ValorPropiedad(c.Name) is decimal)
                    {
                        _ws.Cells[fila, columna].FormatoNumero("0.00");
                    }
                    else if (r.ValorPropiedad(c.Name) is String)
                    {
                        _ws.Cells[fila, columna].FormatoNumero("@");
                    }

                    _ws.Cells[fila, columna].Value = r.ValorPropiedad(c.Name);

                    _ws.Cells[fila, columna].borderTop(ExcelBorderStyle.Thin);
                    _ws.Cells[fila, columna].borderBottom(ExcelBorderStyle.Thin);
                    _ws.Cells[fila, columna].borderRight(ExcelBorderStyle.Thin);
                    _ws.Cells[fila, columna].borderLeft(ExcelBorderStyle.Thin);

                    if ((fila % 2) == 0)
                    {
                        _ws.Cells[fila, columna].Fondo("#BDBDBD");
                    }
                    else
                    {
                        _ws.Cells[fila, columna].Fondo("#F2F2F2");
                    }

                    _ws.Cells[fila, columna].ColorFuente("#000000");

                    //if (_ws.Cells[fila, columna].Value.ToString().Length > _ws.Cells[1, columna].Value.ToString().Length)
                    //{
                    //    if (_ws.Cells[fila, columna].Value.ToString().Length > AnchoCol)
                    //    {
                    //        _ws.Cells[fila, columna].AutoFitColumns();

                    //        AnchoCol = _ws.Cells[fila, columna].Value.ToString().Length;
                    //    }
                    //}

                    columna++;
                }
            }
        }

        public static JsonResult ErrorJson(this string Mensaje)
        {
            HttpContext.Current.Response.StatusCode = (int)System.Net.HttpStatusCode.InternalServerError;
            HttpContext.Current.Response.StatusDescription = Mensaje;
            return new JsonResult { Data = Mensaje, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        public static JsonResult ErrorJson(this Exception ex)
        {
            HttpContext.Current.Response.StatusCode = (int)System.Net.HttpStatusCode.InternalServerError;
            HttpContext.Current.Response.StatusDescription = ex.Message;
            //ErrorSignal.FromCurrentContext().Raise(ex);
            return new JsonResult { Data = ex.Message, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }

        public static TCopiado CopiaPropiedades<TFuente, TCopiado>(this TFuente fuente, TCopiado copia)
        {
            if (fuente == null || copia == null)
                throw new Exception("Las 2 variables tienen que estar instanciadas.");

            PropertyInfo[] proFuente = fuente.GetType().GetProperties();
            foreach (PropertyInfo prop in proFuente)
            {
                if (copia.GetType().GetProperties().Select(x => x.Name.ToLower()).Contains(prop.Name.ToLower()))
                    copia.GetType().GetProperties().Where(x => x.Name.ToLower() == prop.Name.ToLower()).FirstOrDefault().SetValue(copia, prop.GetValue(fuente, null), null);
            }
            return copia;
        }

        public static string ToNumberFormat(this double num)
        {
            return ((double?)num).ToNumberFormat();
        }

        public static string ToNumberFormat(this int? num)
        {
            return ((double?)num).ToNumberFormat();
        }

        public static string ToNumberFormat(this int num)
        {
            return ((double?)num).ToNumberFormat();
        }

        public static string ToNumberFormat(this double? num)
        {
            string format = "{0:0,0}";
            if (num > 100000)
                format = "{0:#,##0, K}";
            if (num > 1000000)
                format = "{0:#,##0, M}";

            return string.Format(format, num).Replace(",", ".");
        }

        public static string RemueveAcentos(this string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC).Replace(",", "");
        }

        public static string ToCapitalize(this string str)
        {
            return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(str.ToLower());
        }

        public static string StripTagsRegex(this string html)
        {
            string res = Regex.Replace(html, "<.*?>", "<br />");
            return res;
        }

        public static string StripTagsRegex(string html, int max)
        {
            string res = Regex.Replace(html, "<.*?>", string.Empty).Replace("&nbsp;", " ").Replace("\r", " ").Replace("\n", " ");
            if (res.Length > max)
                res = res.Substring(0, max);
            return res;
        }

        public static int ToInt(this string cadena)
        {
            int val = 0;
            int.TryParse(cadena, out val);
            return val;
        }

        public static Int64 ToInt64(this string cadena)
        {
            Int64 val = 0;
            Int64.TryParse(cadena, out val);
            return val;
        }

        public static Decimal ToDecimal(this string cadena)
        {
            decimal val = 0;
            decimal.TryParse(cadena, out val);
            return val;
        }

        public static DateTime ToDateTimeCadena(this string cadena)
        {
            DateTime dt = DateTime.ParseExact(cadena, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            return dt;
        }

        public static DateTime ToDateTimeCadena(this string cadena, string formato)
        {
            DateTime dt = new DateTime();
            DateTime.TryParseExact(cadena, formato, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt);
            return dt;
        }

        public static object ValorPropiedad(this object obj, string nomPropiedad)
        {
            if (obj.GetType().GetProperty(nomPropiedad) != null)
                return obj.GetType().GetProperty(nomPropiedad).GetValue(obj, null);
            else
                return null;
        }

        public static void ValorPropiedad(this object obj, string nomPropiedad, object valor)
        {
            obj.GetType().GetProperty(nomPropiedad).SetValue(obj, valor, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.NonPublic, null, null, System.Globalization.CultureInfo.CurrentCulture);
        }

        public static string AcentosZebra(this string cadena)
        {
            var Acentos = new string[] { "á", "é", "í", "ó", "ú", "Á", "É", "Í", "Ó", "Ú" };
            var AcentosZebra = new string[] { @"\A0", @"\82", @"\A1", @"\A2", @"\A3", @"\B5", @"\90", @"\D6", @"\E3", @"\E9" };

            for (int i = 0; i < Acentos.Length; i++)
            {
                cadena = cadena.Replace(Acentos[i], AcentosZebra[i]);
            }
            return cadena;
        }

        public static string AcompletaRemueve(this string text, int tamanio, string caracter = " ", bool izquierda = true)
        {
            char salto = (char)13;
            char tab = (char)10;
            text = text.Replace(salto, ' ').Replace(tab, ' ').Trim();
            if (text.Length > tamanio)
                return text.Substring(0, tamanio);
            else
            {
                if (text.Length == tamanio)
                    return text;
                else
                {
                    int inicio = text.Length;
                    for (int i = inicio; i < tamanio; i++)
                    {
                        if (izquierda)
                            text = caracter + text;
                        else
                            text += caracter;

                    }
                    return text;
                }
            }
        }

        private static object candado = new object();
                
        public static bool IsPropertyExist(dynamic settings, string name)
        {
            if (settings is ExpandoObject)
                return ((IDictionary<string, object>)settings).ContainsKey(name);

            return settings.GetType().GetProperty(name) != null;
        }

        public static string GetMD5(this string str)
        {
         MD5 md5 = MD5CryptoServiceProvider.Create();
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] stream = null;
            StringBuilder sb = new StringBuilder();
            stream = md5.ComputeHash(encoding.GetBytes(str));
            for (int i = 0; i < stream.Length; i++) sb.AppendFormat("{0:x2}", stream[i]);
            return sb.ToString();
        }
    }
}
