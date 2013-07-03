using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml.Drawing;

namespace ActiveConverter
{
    class Model
    {
        const int FILLED_CELLS_IN_ROW = 5;

        internal int headerRowIndex;
        internal int picturesColIndex;
        internal int codesColIndex;
        internal int descColIndex;
        internal int rrpColIndex;
        internal int suppColIndex;
        internal int actualRowIndex;

        internal ExcelWorksheet sheet;
        internal List<DataItem> data;

        internal string storeDir;
        internal string CategoryIds;

        /// <summary>
        /// Konstruktor
        /// </summary>
        /// <param name="sheet"></param>
        public Model(ExcelWorksheet sheet, string catIds)
        {
            this.sheet = sheet;
            this.data = new List<DataItem>();
            this.CategoryIds = catIds;
        }

        /// <summary>
        /// Najdenie hlavickoveho riadka - to je taky, ktory ma vyplnenych aspon 5 prvych stlpcov
        /// </summary>
        /// <param name="sheet"></param>
        public bool FindHeaderIndex()
        {
            var rowIndex = 1;

            while (rowIndex < 100)
            {
                var row = sheet.Row(rowIndex);

                bool ok = true;
                for (int i = 1; i <= FILLED_CELLS_IN_ROW; i++)
                {
                    var val = (sheet.Cells[rowIndex, i].Value ?? string.Empty).ToString();
                    if (string.IsNullOrEmpty(val))
                    {
                        ok = false;
                        break;
                    }
                }

                if (ok)
                {
                    headerRowIndex = rowIndex;
                    return true;    // nasli sme hlavickovy riadok
                }

                rowIndex++;
            }

            return false;
        }

        /// <summary>
        /// Najde indexy relevantnych stlpcov
        /// </summary>
        /// <param name="excelWorksheet"></param>
        internal bool FindColumnIndices()
        {
            if (headerRowIndex < 1)
                return false;

            string val = string.Empty;
            int col = 1;

            do
            {
                val = (sheet.Cells[headerRowIndex, col].Value ?? string.Empty).ToString().ToLower();

                if (!string.IsNullOrEmpty(val))
                {
                    if (val.Contains("picture") && picturesColIndex == 0)
                        picturesColIndex = col;
                    else if (val.Contains("code") && codesColIndex == 0)
                        codesColIndex = col;
                    else if (val.Contains("description") && descColIndex == 0)
                        descColIndex = col;
                    else if (val.Contains("rrp") && rrpColIndex == 0)
                        rrpColIndex = col;
                    else if (val.Contains("supp") && suppColIndex == 0)
                        suppColIndex = col;
                    
                    col++;
                }
            } while (!string.IsNullOrEmpty(val));

            return picturesColIndex != 0 &&
                codesColIndex != 0 &&
                descColIndex != 0 &&
                rrpColIndex != 0 &&
                suppColIndex != 0;
        }

        /// <summary>
        /// Najde najblizsi zaciatok sekcie s udajmi dalsieho produktu
        /// </summary>
        /// <returns></returns>
        internal bool FindFirstCode()
        {
            if (actualRowIndex == 0)
                actualRowIndex = headerRowIndex + 1;

            int attempts = 0;
            while (attempts < 100)  // max 100 prazdnych riadkov a potom to prehlasim za koniec suboru
            {
                var val = (sheet.Cells[actualRowIndex, codesColIndex].Value ?? string.Empty).ToString();
                if (!string.IsNullOrEmpty(val))
                {
                    return true;    // nasli sme kod
                }
                else
                    actualRowIndex++;

                attempts++;
            }

            return false;   // ak nic nenajdeme attempts dosiahnu 100vky a to znamena neuspech
        }

        /// <summary>
        /// Nacita dalsi produkt
        /// </summary>
        /// <returns></returns>
        internal bool ProcessNextDataItem()
        {
            DataItem New = new DataItem();
            string code = string.Empty;

            // actualRowIndex je na prvom zazname.. obrazok by mal tento prekryvat
            object pic = FindPictureForActual();
            if (pic != null)
                New.Picture = pic;

            do
            {
                code = (sheet.Cells[actualRowIndex, codesColIndex].Value ?? string.Empty).ToString();
                if (!string.IsNullOrEmpty(code))
                {
                    // kod polozky
                    New.Codes.Add(code);

                    // popis - nazov a velkost
                    var desc = (sheet.Cells[actualRowIndex, descColIndex].Value ?? string.Empty).ToString();
                    if (string.IsNullOrEmpty(New.Description))
                        New.Description = GetDescription(desc);
                    New.Sizes.Add(GetSize(desc));

                    // rrp
                    var rrp = (sheet.Cells[actualRowIndex, rrpColIndex].Value ?? string.Empty).ToString();
                    if (string.IsNullOrEmpty(New.Rrp))
                        New.Rrp = rrp;

                    // supp
                    var supp = (sheet.Cells[actualRowIndex, suppColIndex].Value ?? string.Empty).ToString();
                    if (string.IsNullOrEmpty(New.Supp))
                        New.Supp = supp;

                    // presun na dalasi riadok
                    actualRowIndex++;
                }
            } while (!string.IsNullOrEmpty(code));

            data.Add(New);

            return true;
        }

        private object FindPictureForActual()
        {
            for (int i = 0; i < sheet.Drawings.Count; i++)
            {
                var d = sheet.Drawings[i];
                if (actualRowIndex <= d.To.Row && actualRowIndex >= d.From.Row)
                    return d;
            }

            return null;
        }

        private string GetSize(string desc)
        {
            if (string.IsNullOrEmpty(desc))
                return string.Empty;

            var index = desc.LastIndexOf('(');
            if (index == -1)
                return desc;

            return desc.Substring(index).Trim('(', ')').Trim();
        }

        private string GetDescription(string desc)
        {
            if (string.IsNullOrEmpty(desc))
                return string.Empty;

            var index = desc.LastIndexOf('(');
            if (index == -1)
                return desc;

            return desc.Substring(0, index).Trim();
        }

        /// <summary>
        /// Ulozenie nacitanych dat do CSV a obrazky zvlast
        /// </summary>
        /// <param name="p"></param>
        internal bool StoreData(string dir)
        {
            if (string.IsNullOrEmpty(dir))
                return false;

            var ts = DateTime.Now.ToString("yyyyMMdd.HH.mm.ss");
            storeDir = dir + (@"\" + ts);

            if (!Directory.Exists(storeDir))
                Directory.CreateDirectory(storeDir);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(GetCSVHeader());
            for (int i = 0; i < data.Count; i++)
            {
                sb.AppendLine(GetNextLine(data[i]));
            }

            File.WriteAllText(storeDir + @"\data.csv", sb.ToString());

            return true;
        }

        private string GetNextLine(DataItem dataItem)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("\"sk\",");//store
            sb.Append(dataItem.Codes[0] + ",");//sku
            sb.Append("\"Default\",");//_attribute_set 
            sb.Append("\"simple\",");//_type 
            sb.Append("\"0\",");//magmi:delete 
            sb.Append("\"1\",");//status 
            sb.Append("\"4\",");//visibility 
            var desc = "\"" + dataItem.Description + "\",";
            sb.Append(desc);//name 
            sb.Append(desc.ToLower().Replace(' ', '-').RemoveDiacritics());//url_key 
            sb.Append("\"\",");//description 
            sb.Append("\"\",");//short_description 
            sb.Append("\"1\",");//enable_googlecheckout  
            sb.Append("\""+CategoryIds+"\",");//category_ids 
            sb.Append("\"1\",");//weight
            sb.Append("\"" + dataItem.Rrp + "\",");//price
            sb.Append("\"" + dataItem.Rrp + "\",");//special_price 
            sb.Append("\"" + DateTime.Now.ToString("yyyy-MM-dd") +"\",");//special_from_date 
            sb.Append("\"2\",");//tax_class_id 
            sb.Append("\"0\",");//manage_stock 
            sb.Append("\"0\",");//use_config_manage_stock 
            sb.Append("\"" + dataItem.Supp + "\",");//brand 
            sb.Append("\"" + string.Join("|", dataItem.Sizes.ToArray()) + "\",");//Veľkosť:Dropdown:1
            
            string picName = StorePix(dataItem);
            var compl = "\"" + picName + "\",";
            sb.Append(compl);//thumbnail 
            sb.Append(compl);//small_image 
            sb.Append(compl);//image 
            sb.Append(compl);//media_gallery 
            sb.Append("\"0\"");//media_gallery_reset 

            return sb.ToString();
        }

        private string StorePix(DataItem dataItem)
        {
            if (dataItem.Picture == null)
                return string.Empty;
            if (!(dataItem.Picture is ExcelPicture))
                return string.Empty;

            var pic = dataItem.Picture as ExcelPicture;
            var name = pic.Name + "." + pic.ImageFormat;
            pic.Image.Save(storeDir + @"\" + name);

            return name;
        }

        private string GetCSVHeader()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("store,");
            sb.Append("sku,");
            sb.Append("_attribute_set,");
            sb.Append("_type,");
            sb.Append("magmi:delete,");
            sb.Append("status,");
            sb.Append("visibility,");
            sb.Append("name,");
            sb.Append("url_key,");
            sb.Append("description,");
            sb.Append("short_description,");
            sb.Append("enable_googlecheckout,");
            sb.Append("category_ids,");
            sb.Append("weight,");
            sb.Append("price,");
            sb.Append("special_price,");
            sb.Append("special_from_date,");
            sb.Append("tax_class_id,");
            sb.Append("manage_stock,");
            sb.Append("use_config_manage_stock,");
            sb.Append("brand,");
            sb.Append("Veľkosť:Dropdown:1,");
            sb.Append("thumbnail,");
            sb.Append("small_image,");
            sb.Append("image,");
            sb.Append("media_gallery,");
            sb.Append("media_gallery_reset");

            return sb.ToString();
        }
    }
}
