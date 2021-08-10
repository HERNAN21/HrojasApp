using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace HrojasApp.Helpers
{
    public class HelperDatos
    {
        public static DataTable XElementToDataTable(List<XElement> listaXElement, string subNivel)
        {
            DataTable dtable = new DataTable();
            if (subNivel == "")
            {
                // construimos columnas del DataTable
                foreach (XElement xe in listaXElement.Descendants())
                    dtable.Columns.Add(new DataColumn(xe.Name.ToString(), typeof(string)));

                DataRow dr = dtable.NewRow();
                foreach (XElement itemElement in listaXElement.Descendants())
                {
                    dr[itemElement.Name.ToString()] = itemElement.Value; //add in the values
                }
                dtable.Rows.Add(dr);
            }
            else
            {
                XElement xelementColumn = (from p in listaXElement.Descendants(subNivel) select p).First();
                // construimos columnas del DataTable
                foreach (XElement xe in xelementColumn.Descendants())
                    dtable.Columns.Add(new DataColumn(xe.Name.ToString(), typeof(string)));

                int ind = 0;
                DataRow dr = null;
                foreach (XElement itemElement in listaXElement.Descendants())
                {
                    if (itemElement.Name.ToString() == subNivel)
                    {
                        if (ind > 0)
                        {
                            dtable.Rows.Add(dr);
                            ind = 0;
                        }
                        dr = dtable.NewRow();
                        ind++;
                    }
                    else
                    {
                        dr[itemElement.Name.ToString()] = itemElement.Value;
                    }
                }
                dtable.Rows.Add(dr);
            }
            dtable.AcceptChanges();
            return dtable;
        }
        public static string DataTableToXMLString(DataTable dt)
        {
            XmlDocument soapEnvelopeXml = null;
            soapEnvelopeXml = new XmlDocument();
            StringBuilder xmlString = new StringBuilder();
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Encoding = Encoding.UTF8,
                ConformanceLevel = ConformanceLevel.Document,
                OmitXmlDeclaration = true,
                CloseOutput = true,
                Indent = true,
                IndentChars = "  ",
                NewLineHandling = NewLineHandling.Replace
            };

            using (XmlWriter xmlWriter = XmlWriter.Create(xmlString, settings))
            {
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("SACV");

                foreach (DataRow dRow in dt.Rows)
                {
                    xmlWriter.WriteStartElement("item");
                    foreach (DataColumn dataColumn in dt.Columns)
                    {
                        xmlWriter.WriteElementString(dataColumn.ColumnName, dRow[dataColumn].ToString());
                    }
                    xmlWriter.WriteEndElement();
                }

                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndDocument();
            }
            string cadenaXml = xmlString.ToString();
            cadenaXml = cadenaXml.Replace("\r\n    ", "").Replace("\r\n   ", "").Replace("\r\n  ", "").Replace("\r\n ", "").Replace("\r\n", "");
            return cadenaXml;
        }
        public static DataTable ToDataTable<T>(IEnumerable<T> collection)
        {
            DataTable dt = new DataTable();
            Type t = typeof(T);
            PropertyInfo[] pia = t.GetProperties();
            object temp;
            DataRow dr;

            for (int i = 0; i < pia.Length; i++)
            {
                dt.Columns.Add(pia[i].Name, Nullable.GetUnderlyingType(pia[i].PropertyType) ?? pia[i].PropertyType);
                dt.Columns[i].AllowDBNull = true;
            }

            //Populate the table
            foreach (T item in collection)
            {
                dr = dt.NewRow();
                dr.BeginEdit();

                for (int i = 0; i < pia.Length; i++)
                {
                    temp = pia[i].GetValue(item, null);
                    if (temp == null || (temp.GetType().Name == "Char" && ((char)temp).Equals('\0')))
                    {
                        dr[pia[i].Name] = (object)DBNull.Value;
                    }
                    else
                    {
                        dr[pia[i].Name] = temp;
                    }
                }

                dr.EndEdit();
                dt.Rows.Add(dr);
            }
            return dt;
        }
        public static DataTable XPDataViewToDataTable(DevExpress.Xpo.XPDataView xpDataView)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            System.Data.DataColumn colunm = null;
            System.Data.DataRow row;
            //  
            foreach (DevExpress.Xpo.DataViewProperty item in xpDataView.Properties)
            {
                colunm = new System.Data.DataColumn(item.Name, item.ValueType);
                colunm.AllowDBNull = true;
                dt.Columns.Add(colunm);
            }

            //  
            foreach (DevExpress.Xpo.DataViewRecord dataRow in xpDataView)
            {
                row = dt.NewRow();
                foreach (DevExpress.Xpo.DataViewProperty dataColunm in xpDataView.Properties)
                {
                    if (dataRow[dataColunm.Name] == null)
                    {
                        row[dataColunm.Name] = System.DBNull.Value;
                    }
                    else
                    {
                        row[dataColunm.Name] = dataRow[dataColunm.Name];
                    }
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        public static DevExpress.Xpo.XPDataView SelectedDataToXPDataView(DevExpress.Xpo.DB.SelectedData selected)
        {
            DevExpress.Xpo.XPDataView dvResultado = new DevExpress.Xpo.XPDataView();
            foreach (var row in selected.ResultSet[0].Rows)
            {
                dvResultado.AddProperty((string)row.Values[0], DevExpress.Xpo.DB.DBColumn.GetType((DevExpress.Xpo.DB.DBColumnType)Enum.Parse(typeof(DevExpress.Xpo.DB.DBColumnType), (string)row.Values[2])));
            }
            dvResultado.LoadData(new DevExpress.Xpo.DB.SelectedData(selected.ResultSet[1]));
            return dvResultado;
        }
        public static string ColumnaSeleccionada(int num)
        {
            switch (num)
            {
                case 1:
                    return "A";
                case 2:
                    return "B";
                case 3:
                    return "C";
                case 4:
                    return "D";
                case 5:
                    return "E";
                case 6:
                    return "F";
                case 7:
                    return "G";
                case 8:
                    return "H";
                case 9:
                    return "I";
                case 10:
                    return "J";
                case 11:
                    return "K";
                case 12:
                    return "L";
                case 13:
                    return "M";
                case 14:
                    return "N";
                case 15:
                    return "O";
                case 16:
                    return "P";
                case 17:
                    return "Q";
                case 18:
                    return "R";
                case 19:
                    return "S";
                case 20:
                    return "T";
                case 21:
                    return "U";
                case 22:
                    return "V";
                case 23:
                    return "W";
                case 24:
                    return "X";
                case 25:
                    return "Y";
                case 26:
                    return "Z";
                case 27:
                    return "AA";
                case 28:
                    return "AB";
                case 29:
                    return "AC";
                case 30:
                    return "AD";
                case 31:
                    return "AE";
                case 32:
                    return "AF";
                case 33:
                    return "AG";
                case 34:
                    return "AH";
                case 35:
                    return "AI";
                case 36:
                    return "AJ";
                case 37:
                    return "AK";
                case 38:
                    return "AL";
                case 39:
                    return "AM";
                case 40:
                    return "AN";
                case 41:
                    return "AO";
                case 42:
                    return "AP";
                case 43:
                    return "AQ";
                case 44:
                    return "AR";
                case 45:
                    return "AS";
                case 46:
                    return "AT";
                case 47:
                    return "AU";
                case 48:
                    return "AV";
                case 49:
                    return "AW";
                case 50:
                    return "AX";
                case 51:
                    return "AY";
                case 52:
                    return "AZ";
                case 53:
                    return "BA";
                case 54:
                    return "BB";
                case 55:
                    return "BC";
                case 56:
                    return "BD";
                case 57:
                    return "BE";
                case 58:
                    return "BF";
                case 59:
                    return "BG";
                case 60:
                    return "BH";
                default:
                    return "BI";
            }
        }
        public static List<string> SepararNombresApellidos(string NombreCompleto)
        {
            List<string> ListaApe = new List<string>();
            int contador = 0;
            string nom = string.Empty;
            string[] split = NombreCompleto.Split(new Char[] { ' ' });
            foreach (string s in split)
            {
                if (contador < 2)
                    ListaApe.Add(s);
                else if (contador == 2)
                    nom += s;
                else
                    nom += " " + s;
                contador++;
            }
            ListaApe.Add(nom);
            return ListaApe;
        }
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }
    }
}
