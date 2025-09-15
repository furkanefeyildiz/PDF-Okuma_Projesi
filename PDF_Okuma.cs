using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Globalization;

namespace PDF_OkumaProjesi
{

    public partial class PDF_Okuma : Form
    {
        public PDF_Okuma()
        {
            InitializeComponent();
        }

        OracleConnection con = new OracleConnection(@"Data Source=10.219.222.38:1521/MARSHDB;Persist Security Info=True;User ID=RAPOR;Password=M2cY_WeL3cAt4_2o25k;");
        private void btnSelectPDF_Click(object sender, EventArgs e)
        {

            if (cmbBelgeTuru.SelectedItem == null)
            {
                MessageBox.Show("Lütfen Belge Türü Seçin");
            }
            else if (cmbSirketSecimi.SelectedItem == null)
            {
                MessageBox.Show("Lütfen Şirket Seçin");
            }
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "PDF Dosyaları|*.pdf";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfPath = openFileDialog.FileName;
                    string content = ReadPdfText(pdfPath);
                    txtPDFContent.Text = content;    //PDF İÇERİĞİ
                    MessageBox.Show("PDF Yüklendi");
                    btnExportExcel.Enabled = true;
                }
            }


        }

        private string ReadPdfText(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                    text.Append(Environment.NewLine);
                }
                return text.ToString();
            }

        }
        private List<(string Teminat, string Bedel)> ExtractTeminatlar(string pdfText)
        {
            var result = new List<(string Teminat, string Bedel)>();
            var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim())
            .ToList();

            bool inTable = false;
            string bufferTeminat = null;

            var lineWithBedelPattern = new Regex(@"^(?<left>.+?)\s+(?<bedel>(?:RAYİÇ\s*BEDEL|DAHİL|\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?))\s*$", RegexOptions.IgnoreCase);

            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();

                if (!inTable)
                {
                    if (line.ToUpper().Contains("TEMİNATLAR") || line.ToUpper().Contains("SİGORTA BEDELİ") || line.ToUpper().Contains("6ù*257$7(0ù1$7,") || line.ToUpper().Contains("BEDEL (EUR)"))
                    {
                        inTable = true;
                    }
                    continue;
                }

                if (line.Contains("Ferdi Kaza"))
                    break;

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                if (line.ToUpper().StartsWith("ARAÇTA BULUNAN K. EŞYA") || line.ToUpper().StartsWith("HIRSIZLIK(DEMİRBAŞ)") || line.ToUpper().StartsWith("HUKUKSAL KORUMA") || line.ToUpper().StartsWith(": Poliçe No 772287393 /"))
                {
                    var l = lineWithBedelPattern.Match(line);
                    if (l.Success)
                    {
                        result.Add((l.Groups["left"].Value.Trim(), l.Groups["bedel"].Value.Trim()));
                    }
                    else
                    {
                        result.Add((line, ""));
                    }
                    break;
                }

                if (Regex.IsMatch(line, @"^(DAHİL|RAYİÇ\s*BEDEL|\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)$", RegexOptions.IgnoreCase))
                {
                    if (!string.IsNullOrEmpty(bufferTeminat))
                    {
                        result.Add((bufferTeminat.Trim(), line));
                        bufferTeminat = null;
                    }
                    continue;
                }

                var m = lineWithBedelPattern.Match(line);
                if (m.Success)
                {
                    string teminat = m.Groups["left"].Value.Trim();
                    string bedel = m.Groups["bedel"].Value.Trim();

                    if (!string.IsNullOrEmpty(bufferTeminat))
                    {
                        teminat = (bufferTeminat + " " + teminat).Trim();
                        bufferTeminat = null;
                    }

                    result.Add((teminat, bedel));
                }
                else
                {
                    if (bufferTeminat == null)
                        bufferTeminat = line;
                    else
                        bufferTeminat += " " + line;
                }
            }

            if (!string.IsNullOrEmpty(bufferTeminat))
            {
                result.Add((bufferTeminat.Trim(), ""));
            }

            return result;
        }

        private List<(DateTime Tarih, decimal Tutar)> GetTaksitler(string pdfText)
        {
            var taksitler = new List<(DateTime, decimal)>();
            var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                // Regex: Tarih + boşluk + tutar (ondalık virgülle)
                var match = Regex.Match(trimmedLine, @"(\d{2}/\d{2}/\d{4})\s+(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)");

                if (match.Success)
                {
                    string tarihStr = match.Groups[1].Value;
                    string tutarStr = match.Groups[2].Value;

                    // Tarihi parse et
                    if (DateTime.TryParseExact(tarihStr, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime tarih))
                    {
                        // Tutarı normalize et
                        string normalizedTutar = tutarStr.Replace(".", "").Replace(',', '.');

                        if (decimal.TryParse(normalizedTutar, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tutar))
                        {
                            if (tutar < 21)
                                continue;

                            taksitler.Add((tarih, tutar));
                        }
                    }
                }
            }

            return taksitler;
        }





        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            string pdfText = txtPDFContent.Text;

            if (cmbSirketSecimi.Text == "AXA" && pdfText.ToUpper().Contains("AXA SIGORTA A.S."))
            {
                //AXA SİGORTA

                string poliçeNo = "";
                string musteriNo = "";
                string baslangicTarihi = "";
                string bitisTarihi = "";
                string tanzimTarihi = "";
                string taksitler = "";
                string ekBelgeNo = "";
                var taksitlerTuple = GetTaksitler(pdfText);
                var taksitList = new List<string>();
                var taksitTarihleri = new List<DateTime>();


                var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    var matchPoliçe = Regex.Match(line, @"Poliçe No\s*:*\s*:*\s*([0-9]+)");
                    if (matchPoliçe.Success)
                        poliçeNo = matchPoliçe.Groups[1].Value.Trim();

                    var matchMusteri = Regex.Match(line, @"0ûWHUL1R \s*:*\s*:*\s*([0-9]+)");
                    if (matchMusteri.Success)
                        musteriNo = matchMusteri.Groups[1].Value.Trim();

                    var matchBaslangic = Regex.Match(line, @"%DûODQJÖo7DULKL \s*:*\s*:*\s*([0-9]{2}/[0-9]{2}/[0-9]{4})");
                    if (matchBaslangic.Success)
                        baslangicTarihi = matchBaslangic.Groups[1].Value.Trim();

                    var matchBitis = Regex.Match(line, @"%LWLû7DULKL \s*:*\s*:*\s*([0-9]{2}/[0-9]{2}/[0-9]{4})");
                    if (matchBitis.Success)
                        bitisTarihi = matchBitis.Groups[1].Value.Trim();

                    var matchTanzim = Regex.Match(line, @"Tanzim Tarihi\s*:*\s*:*\s*([0-9]{2}/[0-9]{2}/[0-9]{4})");
                    if (matchTanzim.Success)
                        tanzimTarihi = matchTanzim.Groups[1].Value.Trim();

                    var matchEkBelge = Regex.Match(line, @"Ek Belge No\s*:*\s*([0-9]+)");
                    if (matchEkBelge.Success)
                        ekBelgeNo = matchEkBelge.Groups[1].Value.Trim();
                }



                foreach (var (tarih, tutar) in taksitlerTuple)
                {
                    taksitTarihleri.Add(tarih);
                    taksitList.Add(tutar.ToString(CultureInfo.InvariantCulture));
                }

                taksitler = string.Join("\n", taksitList);



                var teminatlar = ExtractTeminatlar(pdfText);

                // TEMİNATLARI AYRI EXCEL'E YAZ
                string saveDirectory = @"C:\EXCEL";
                if (!Directory.Exists(saveDirectory))
                    Directory.CreateDirectory(saveDirectory);

                string baseFileName = "axa_teminatlar.xlsx";
                string excelPath = GetUniqueFileName(saveDirectory, baseFileName);

                using (var excel = new OfficeOpenXml.ExcelPackage())
                {
                    var ws = excel.Workbook.Worksheets.Add("Teminatlar");

                    ws.Cells["C1"].Value = "Poliçe No";
                    ws.Cells["C2"].Value = poliçeNo;

                    ws.Cells["A1"].Value = "Teminat";
                    ws.Cells["B1"].Value = "Sigorta Bedeli";

                    int row = 2;
                    foreach (var t in teminatlar)
                    {
                        ws.Cells[row, 1].Value = t.Teminat;
                        ws.Cells[row, 2].Value = t.Bedel;
                        row++;
                    }

                    ws.Cells.AutoFitColumns();
                    excel.SaveAs(new FileInfo(excelPath));
                }

                SaveToExcel(poliçeNo, musteriNo, baslangicTarihi, bitisTarihi, tanzimTarihi, taksitler, ekBelgeNo, taksitTarihleri);

                //MessageBox.Show($"Teminatlar Excel'e yazıldı: {excelPath}", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                string firmCode = "2";
                string companyName = cmbSirketSecimi.Text.Trim();
                string productName;
                if (pdfText.Contains("(1'h675ù<("))
                    productName = "ENDÜSTRİYEL PAKET YANGIN POLİÇESİ";
                else
                {
                    productName = cmbBelgeTuru.Text.Trim();
                }
                string curType = "TRY";



                try
                {
                    con.Open();


                    //Insert POLICYMASTER
                    OracleCommand cmd1 = new OracleCommand(@"
        INSERT INTO mercury2011.Z_PDFREADER_POLICYMASTER
        (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, CLIENT_NO, BEG_DATE, END_DATE, CONFIRM_DATE, CUR_TYPE)
        VALUES (:firm, :comp, :prod, :pol, :end, :cli, TO_DATE(:beg, 'DD/MM/YYYY'), TO_DATE(:fin, 'DD/MM/YYYY'), TO_DATE(:con, 'DD/MM/YYYY'), :cur)", con);

                    cmd1.Parameters.Add(":firm", firmCode);
                    cmd1.Parameters.Add(":comp", companyName);
                    cmd1.Parameters.Add(":prod", productName);
                    cmd1.Parameters.Add(":pol", poliçeNo);
                    cmd1.Parameters.Add(":end", ekBelgeNo);
                    cmd1.Parameters.Add(":cli", musteriNo);
                    cmd1.Parameters.Add(":beg", baslangicTarihi);
                    cmd1.Parameters.Add(":fin", bitisTarihi);
                    cmd1.Parameters.Add(":con", tanzimTarihi);
                    cmd1.Parameters.Add(":cur", curType);

                    cmd1.ExecuteNonQuery();

                    string[] taksitArr = taksitler.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < taksitArr.Length; i++)
                    {
                        string tutarStr = taksitArr[i]
                            .Replace("USD", "")
                            .Replace("TL", "")
                            .Replace("EUR", "");


                        if (decimal.TryParse(tutarStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tutar))
                        {
                            OracleCommand cmd2 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYINSTALLMENT
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, INSTALLMENT_ORDER, INSTALLMENT_DATE, INSTALLMENT_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :ord, :installdate, :amt)", con);

                            cmd2.Parameters.Add(":firm", firmCode);
                            cmd2.Parameters.Add(":comp", companyName);
                            cmd2.Parameters.Add(":prod", productName);
                            cmd2.Parameters.Add(":pol", poliçeNo);
                            cmd2.Parameters.Add(":end", ekBelgeNo);
                            cmd2.Parameters.Add(":ord", i + 1);
                            cmd2.Parameters.Add(":installdate", taksitTarihleri[i].ToString("dd/MM/yyyy"));
                            cmd2.Parameters.Add(":amt", tutar);

                            cmd2.ExecuteNonQuery();
                        }
                        else
                        {
                            MessageBox.Show($"Taksit tutarı geçersiz formatta: {taksitArr[i]}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }

                    //Insert POLICYCOVERAGE
                    foreach (var t in teminatlar)
                    {
                        string bedelStr = t.Bedel.Replace("TL", "").Replace("USD", "").Replace("EUR", "").Trim();
                        if (string.IsNullOrEmpty(bedelStr)) bedelStr = "0";

                        OracleCommand cmd3 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYCOVERAGE
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, COVERAGE_NAME, COVERAGE_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :cov, :amt)", con);

                        cmd3.Parameters.Add(new OracleParameter(":firm", firmCode));
                        cmd3.Parameters.Add(new OracleParameter(":comp", companyName));
                        cmd3.Parameters.Add(new OracleParameter(":prod", productName));
                        cmd3.Parameters.Add(new OracleParameter(":pol", poliçeNo));
                        cmd3.Parameters.Add(new OracleParameter(":end", ekBelgeNo));
                        cmd3.Parameters.Add(new OracleParameter(":cov", t.Teminat));
                        cmd3.Parameters.Add(new OracleParameter(":amt", bedelStr));

                        cmd3.ExecuteNonQuery();
                    }

                    MessageBox.Show("SQL'e başarıyla kaydedildi.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("SQL Hatası: " + ex.Message);
                }
                finally
                {
                    if (con.State == System.Data.ConnectionState.Open)
                        con.Close();
                }
            }



            else if (cmbSirketSecimi.Text == "TURKIYE" && pdfText.ToUpper().Contains("TÜRKİYE SİGORTA AŞ"))
            {
                // TÜRKİYE SİGORTA

                string poliçeNo = "";
                string musteriNo = "";
                string baslangicTarihi = "";
                string bitisTarihi = "";
                string tanzimTarihi = "";
                string taksitler = "";
                string ekBelgeNo = "";

                var taksitTarihleri = new List<DateTime>();

                var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                // ---------------- POLİÇE BİLGİLERİ ----------------
                for (int i = 0; i < lines.Length - 1; i++)
                {
                    if (Regex.IsMatch(lines[i], @"^\d{2}\.\d{2}\.\d{4}$"))
                    {
                        tanzimTarihi = lines[i].Trim();
                    }

                    if (Regex.IsMatch(lines[i], @"^\d{6,}\s+\d+\s+\d+\/\d+"))
                    {
                        var values = Regex.Split(lines[i].Trim(), @"\s+");

                        if (values.Length >= 7)
                        {
                            musteriNo = values[0];
                            poliçeNo = values[2];
                            ekBelgeNo = values[3];
                            baslangicTarihi = values[4];
                            bitisTarihi = values[5];
                        }

                        break;
                    }
                }

                // ---------------- TAKSİTLER ----------------
                List<string> taksitlerList = new List<string>();

                foreach (var line in lines)
                {
                    // Sadece "Peşinat 27.10.2024 21.809,56 TL" şeklinde satırları yakala
                    var match = Regex.Match(line, @"^(Peşinat|\d+)\s+(\d{2}\.\d{2}\.\d{4})\s+([\d\.,]+)\s*TL");

                    if (match.Success)
                    {
                        string tarihStr = match.Groups[2].Value;
                        string tutarStr = match.Groups[3].Value;

                        // Sadece tutarı ekle (TL ile birlikte)
                        taksitlerList.Add(tutarStr);

                        // Tarihi DateTime'a çevir
                        if (DateTime.TryParseExact(tarihStr, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime tarih))
                        {
                            taksitTarihleri.Add(tarih);
                        }
                    }
                }

                taksitler = string.Join("\n", taksitlerList);
                // ---------------------------------------------

                var teminatlar = ExtractTeminatlar(pdfText);

                // ---------------- EXCEL ----------------
                string saveDirectory = @"C:\EXCEL";
                if (!Directory.Exists(saveDirectory))
                    Directory.CreateDirectory(saveDirectory);

                string baseFileName = "turkiye_teminatlar.xlsx";
                string excelPath = GetUniqueFileName(saveDirectory, baseFileName);

                using (var excel = new OfficeOpenXml.ExcelPackage())
                {
                    var ws = excel.Workbook.Worksheets.Add("Teminatlar");

                    ws.Cells[1, 1].Value = "Poliçe No";
                    ws.Cells[2, 1].Value = poliçeNo;

                    int col = 2;
                    foreach (var t in teminatlar)
                    {
                        ws.Cells[1, col].Value = t.Teminat;
                        ws.Cells[2, col].Value = t.Bedel;
                        col++;
                    }

                    ws.Cells.AutoFitColumns();
                    excel.SaveAs(new FileInfo(excelPath));
                }

                // Excel'e poliçe bilgilerini + taksitleri yaz
                SaveToExcel(poliçeNo, musteriNo, baslangicTarihi, bitisTarihi, tanzimTarihi, taksitler, ekBelgeNo, taksitTarihleri);



                //MessageBox.Show($"Teminatlar Excel'e yazıldı: {excelPath}", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                string firmCode = "2";
                string companyName = cmbSirketSecimi.Text.Trim();
                string productName = "TRAFİK SİGORTA POLİÇESİ";
                string curType = "TRY";

                try
                {
                    con.Open();


                    //Insert POLICYMASTER
                    OracleCommand cmd1 = new OracleCommand(@"
        INSERT INTO mercury2011.Z_PDFREADER_POLICYMASTER
        (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, CLIENT_NO, BEG_DATE, END_DATE, CONFIRM_DATE, CUR_TYPE)
        VALUES (:firm, :comp, :prod, :pol, :end, :cli, TO_DATE(:beg, 'DD/MM/YYYY'), TO_DATE(:fin, 'DD/MM/YYYY'), TO_DATE(:con, 'DD/MM/YYYY'), :cur)", con);

                    cmd1.Parameters.Add(":firm", firmCode.Trim());
                    cmd1.Parameters.Add(":comp", companyName.Trim());
                    cmd1.Parameters.Add(":prod", productName.Trim());
                    cmd1.Parameters.Add(":pol", poliçeNo.Trim());
                    cmd1.Parameters.Add(":end", ekBelgeNo.Trim());
                    cmd1.Parameters.Add(":cli", musteriNo.Trim());
                    cmd1.Parameters.Add(":beg", baslangicTarihi.Trim());
                    cmd1.Parameters.Add(":fin", bitisTarihi.Trim());
                    cmd1.Parameters.Add(":con", tanzimTarihi.Trim());
                    cmd1.Parameters.Add(":cur", curType.Trim());

                    cmd1.ExecuteNonQuery();

                    string[] taksitArr = taksitler.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < taksitArr.Length; i++)
                    {
                        string tutarStr = taksitArr[i]
                            .Replace("USD", "")
                            .Replace("TL", "")
                            .Replace("EUR", "");


                        if (decimal.TryParse(tutarStr, NumberStyles.Any, new CultureInfo("tr-TR"), out decimal tutar))
                        {
                            OracleCommand cmd2 = new OracleCommand(@"
    INSERT INTO mercury2011.Z_PDFREADER_POLICYINSTALLMENT
    (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, INSTALLMENT_ORDER, INSTALLMENT_DATE, INSTALLMENT_AMOUNT)
    VALUES (:firm, :comp, :prod, :pol, :endorsNo, :ord, :instDate, :amt)", con);

                            cmd2.Parameters.Add(":firm", firmCode);
                            cmd2.Parameters.Add(":comp", companyName);
                            cmd2.Parameters.Add(":prod", productName);
                            cmd2.Parameters.Add(":pol", poliçeNo);
                            cmd2.Parameters.Add(":endorsNo", ekBelgeNo);
                            cmd2.Parameters.Add(":ord", i + 1);
                            cmd2.Parameters.Add(":instDate", taksitTarihleri[i]);
                            cmd2.Parameters.Add(":amt", tutar);

                            cmd2.ExecuteNonQuery();
                        }
                        else
                        {
                            MessageBox.Show($"Taksit tutarı geçersiz formatta: {taksitArr[i]}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }

                    //Insert POLICYCOVERAGE
                    foreach (var t in teminatlar)
                    {
                        string bedelStr = t.Bedel.Replace("TL", "").Replace("USD", "").Replace("EUR", "").Trim();
                        if (string.IsNullOrEmpty(bedelStr)) bedelStr = "0";

                        OracleCommand cmd3 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYCOVERAGE
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, COVERAGE_NAME, COVERAGE_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :cov, :amt)", con);

                        cmd3.Parameters.Add(new OracleParameter(":firm", firmCode));
                        cmd3.Parameters.Add(new OracleParameter(":comp", companyName));
                        cmd3.Parameters.Add(new OracleParameter(":prod", productName));
                        cmd3.Parameters.Add(new OracleParameter(":pol", poliçeNo));
                        cmd3.Parameters.Add(new OracleParameter(":end", ekBelgeNo));
                        cmd3.Parameters.Add(new OracleParameter(":cov", t.Teminat));
                        cmd3.Parameters.Add(new OracleParameter(":amt", bedelStr));

                        cmd3.ExecuteNonQuery();
                    }

                    MessageBox.Show("SQL'e başarıyla kaydedildi.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("SQL Hatası: " + ex.Message);
                }
                finally
                {
                    if (con.State == System.Data.ConnectionState.Open)
                        con.Close();
                }


            }


            //SOMPO SİGORTA
            if (cmbSirketSecimi.Text == "SOMPO")
            {
                if (pdfText.ToUpper().Contains("SOMPO SİGORTA") && pdfText.ToUpper().Contains("GENİŞLETİLMİŞ FULL KASKO SİGORTA POLİÇESİ"))
                {
                    string poliçeNo = "";
                    string musteriNo = "";
                    string baslangicTarihi = "";
                    string bitisTarihi = "";
                    string tanzimTarihi = "";
                    string taksitler = "";
                    string ekBelgeNo = "";
                    var taksitTarihleri = new List<DateTime>();

                    var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var line in lines)
                    {
                        var values = Regex.Split(line.Trim(), @"\s+");

                        if (values.Length >= 8 &&
                        Regex.IsMatch(values[0], @"^\d{6,}$") && // Müşteri No
                        Regex.IsMatch(values[2], @"^\d{15}$") && // Poliçe No
                        Regex.IsMatch(values[5], @"^\d{2}/\d{2}/\d{4}$") && // Başlangıç
                        Regex.IsMatch(values[6], @"^\d{2}/\d{2}/\d{4}$")) // Bitiş
                        {
                            musteriNo = values[0];
                            poliçeNo = values[2];
                            baslangicTarihi = values[5];
                            bitisTarihi = values[6];
                            tanzimTarihi = baslangicTarihi;

                            if (values.Length > 3 && Regex.IsMatch(values[3], @"^\d+$"))
                            {
                                ekBelgeNo = values[3];
                            }
                            break;
                        }
                    }

                    var matchTaksit = Regex.Matches(pdfText,
                        @"(PESIN|TAKSIT\s*\d+)?\s*(\d{2}/\d{2}/\d{4})\s+([\d.,]+)",
                        RegexOptions.IgnoreCase);

                    foreach (Match m in matchTaksit)
                    {
                        string tarihStr = m.Groups[2].Value.Trim();  //  dd/MM/yyyy
                        string tutarStr = m.Groups[3].Value.Trim();  // 123,45

                        if (decimal.TryParse(tutarStr, out decimal tutar))
                        {
                            if(tutar >= 1 && tutar <= 31)
                            {
                                continue;
                            }
                        }

                        taksitler += tutarStr + "\n";

                        if (DateTime.TryParseExact(tarihStr, "dd/MM/yyyy",
                            new CultureInfo("tr-TR"), DateTimeStyles.None, out DateTime tarih))
                        {
                            taksitTarihleri.Add(tarih);
                        }
                    }




                    //TEMİNAT / SİGORTA BEDELİ TABLOSUNU ÇEK
                    var teminatlar = ExtractTeminatlar(pdfText);

                    string saveDirectory = @"C:\EXCEL";
                    if (!Directory.Exists(saveDirectory))
                        Directory.CreateDirectory(saveDirectory);

                    string baseFileName = "sompo_teminatlar.xlsx";
                    string excelPath = GetUniqueFileName(saveDirectory, baseFileName);

                    using (var excel = new OfficeOpenXml.ExcelPackage())
                    {
                        var ws = excel.Workbook.Worksheets.Add("Teminatlar");

                        ws.Cells["C1"].Value = "Poliçe No";
                        ws.Cells["C2"].Value = poliçeNo;

                        ws.Cells["A1"].Value = "Teminat";
                        ws.Cells["B1"].Value = "Sigorta Bedeli";

                        int row = 2;
                        foreach (var t in teminatlar)
                        {
                            ws.Cells[row, 1].Value = t.Teminat;
                            ws.Cells[row, 2].Value = t.Bedel;
                            row++;
                        }

                        ws.Cells.AutoFitColumns();
                        excel.SaveAs(new FileInfo(excelPath));
                    }

                    SaveToExcel(poliçeNo, musteriNo, baslangicTarihi, bitisTarihi, tanzimTarihi, taksitler, ekBelgeNo, taksitTarihleri);

                    //MessageBox.Show($"Teminatlar Excel'e yazıldı: {excelPath}", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    string firmCode = "2";
                    string companyName = cmbSirketSecimi.Text.Trim();
                    string productName = "GENİŞLETİLMİŞ FULL KASKO SİGORTA POLİÇESİ";
                    string curType = "TRY";

                    try
                    {
                        con.Open();


                        //Insert POLICYMASTER
                        OracleCommand cmd1 = new OracleCommand(@"
        INSERT INTO mercury2011.Z_PDFREADER_POLICYMASTER
        (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, CLIENT_NO, BEG_DATE, END_DATE, CONFIRM_DATE, CUR_TYPE)
        VALUES (:firm, :comp, :prod, :pol, :end, :cli, TO_DATE(:beg, 'DD/MM/YYYY'), TO_DATE(:fin, 'DD/MM/YYYY'), TO_DATE(:con, 'DD/MM/YYYY'), :cur)", con);

                        cmd1.Parameters.Add(":firm", firmCode);
                        cmd1.Parameters.Add(":comp", companyName);
                        cmd1.Parameters.Add(":prod", productName);
                        cmd1.Parameters.Add(":pol", poliçeNo);
                        cmd1.Parameters.Add(":end", ekBelgeNo);
                        cmd1.Parameters.Add(":cli", musteriNo);
                        cmd1.Parameters.Add(":beg", baslangicTarihi);
                        cmd1.Parameters.Add(":fin", bitisTarihi);
                        cmd1.Parameters.Add(":con", tanzimTarihi);
                        cmd1.Parameters.Add(":cur", curType);

                        cmd1.ExecuteNonQuery();

                        string[] taksitArr = taksitler.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < taksitArr.Length; i++)
                        {
                            string tutarStr = taksitArr[i]
                                .Replace("USD", "")
                                .Replace("TL", "")
                                .Replace("EUR", "");


                            if (decimal.TryParse(tutarStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tutar))
                            {
                                OracleCommand cmd2 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYINSTALLMENT
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, INSTALLMENT_ORDER, INSTALLMENT_DATE, INSTALLMENT_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :ord, :installdate, :amt)", con);

                                cmd2.Parameters.Add(":firm", firmCode);
                                cmd2.Parameters.Add(":comp", companyName);
                                cmd2.Parameters.Add(":prod", productName);
                                cmd2.Parameters.Add(":pol", poliçeNo);
                                cmd2.Parameters.Add(":end", ekBelgeNo);
                                cmd2.Parameters.Add(":ord", i + 1);
                                cmd2.Parameters.Add(":installdate", taksitTarihleri[i]);
                                cmd2.Parameters.Add(":amt", tutar);

                                cmd2.ExecuteNonQuery();
                            }
                            else
                            {
                                MessageBox.Show($"Taksit tutarı geçersiz formatta: {taksitArr[i]}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }

                        //Insert POLICYCOVERAGE
                        foreach (var t in teminatlar)
                        {
                            string bedelStr = t.Bedel.Replace("TL", "").Replace("USD", "").Replace("EUR", "").Trim();
                            if (string.IsNullOrEmpty(bedelStr)) bedelStr = "0";

                            OracleCommand cmd3 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYCOVERAGE
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, COVERAGE_NAME, COVERAGE_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :cov, :amt)", con);

                            cmd3.Parameters.Add(new OracleParameter(":firm", firmCode));
                            cmd3.Parameters.Add(new OracleParameter(":comp", companyName));
                            cmd3.Parameters.Add(new OracleParameter(":prod", productName));
                            cmd3.Parameters.Add(new OracleParameter(":pol", poliçeNo));
                            cmd3.Parameters.Add(new OracleParameter(":end", ekBelgeNo));
                            cmd3.Parameters.Add(new OracleParameter(":cov", t.Teminat));
                            cmd3.Parameters.Add(new OracleParameter(":amt", bedelStr));

                            cmd3.ExecuteNonQuery();
                        }

                        MessageBox.Show("SQL'e başarıyla kaydedildi.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("SQL Hatası: " + ex.Message);
                    }
                    finally
                    {
                        if (con.State == System.Data.ConnectionState.Open)
                            con.Close();
                    }


                }

                else if (pdfText.ToUpper().Contains("BİLEŞİK ÜRÜN"))
                {
                    string poliçeNo = "";
                    string musteriNo = "";
                    string baslangicTarihi = "";
                    string bitisTarihi = "";
                    string tanzimTarihi = "";
                    string taksitler = "";
                    string ekBelgeNo = "";
                    var taksitTarihleri = new List<DateTime>();

                    var lines = pdfText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var line in lines)
                    {
                        if (string.IsNullOrEmpty(poliçeNo))
                        {
                            var m = Regex.Match(line, @"\b\d{15}\b");
                            if (m.Success)
                                poliçeNo = m.Value.Trim();
                        }

                        if (string.IsNullOrEmpty(musteriNo))
                        {
                            var parts = Regex.Split(line.Trim(), @"\s+");
                            if (parts.Length >= 1 && Regex.IsMatch(parts[0], @"^\d{6,}$"))
                                musteriNo = parts[0];
                        }


                        if (string.IsNullOrEmpty(baslangicTarihi) || string.IsNullOrEmpty(bitisTarihi))
                        {
                            var m = Regex.Match(line, @"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})");
                            if (m.Success)
                            {
                                baslangicTarihi = m.Groups[1].Value.Trim();
                                bitisTarihi = m.Groups[2].Value.Trim();
                                tanzimTarihi = m.Groups[2].Value.Trim();
                            }
                        }

                        if (string.IsNullOrEmpty(ekBelgeNo))
                        {
                            var parts = Regex.Split(line.Trim(), @"\s+");
                            if (parts.Length >= 4 && Regex.IsMatch(parts[3], @"^\d+$"))
                            {
                                ekBelgeNo = parts[3];
                            }
                        }
                    }


                    var taksitList = new List<string>();

                    foreach (var line in lines)
                    {
                        var m = Regex.Match(line.Trim(),
                            @"^(?:PESIN|TAKSIT\s*\d+)?\s*(\d{2}/\d{2}/\d{4})\s+((?:\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?))",
                            RegexOptions.IgnoreCase);

                        if (m.Success)
                        {
                            string tarihStr = m.Groups[1].Value; // tarih
                            string raw = m.Groups[2].Value;      // tutar

                            // --- Tutar normalize ---
                            string normalized;
                            if (raw.Contains('.') && raw.Contains(','))
                            {
                                if (raw.LastIndexOf(',') > raw.LastIndexOf('.'))
                                    normalized = raw.Replace(".", "").Replace(',', '.'); // virgül ondalık
                                else
                                    normalized = raw.Replace(",", ""); // nokta ondalık
                            }
                            else if (raw.Contains(','))
                            {
                                normalized = raw.Replace(".", "").Replace(',', '.'); // virgül ondalık
                            }
                            else
                            {
                                normalized = raw.Replace(",", ""); // sadece nokta varsa
                            }

                            // taksit tutarı listesi
                            taksitList.Add($"{normalized} USD");

                            // taksit tarihi listesi
                            if (DateTime.TryParseExact(tarihStr, "dd/MM/yyyy", new CultureInfo("tr-TR"), DateTimeStyles.None, out DateTime tarih))
                            {
                                taksitTarihleri.Add(tarih);
                            }
                        }

                    }

                    if (taksitList.Count > 0)
                        taksitler = string.Join("\n", taksitList);

                    //TEMİNAT / SİGORTA BEDELİ TABLOSUNU ÇEK
                    var teminatlar = ExtractTeminatlar(pdfText);

                    string saveDirectory = @"C:\EXCEL";
                    if (!Directory.Exists(saveDirectory))
                        Directory.CreateDirectory(saveDirectory);

                    string baseFileName = "sompo_teminatlar.xlsx";
                    string excelPath = GetUniqueFileName(saveDirectory, baseFileName);

                    using (var excel = new OfficeOpenXml.ExcelPackage())
                    {
                        var ws = excel.Workbook.Worksheets.Add("Teminatlar");

                        ws.Cells["C1"].Value = "Poliçe No";
                        ws.Cells["C2"].Value = poliçeNo;

                        ws.Cells["A1"].Value = "Teminat";
                        ws.Cells["B1"].Value = "Sigorta Bedeli";

                        int row = 2;
                        foreach (var t in teminatlar)
                        {
                            ws.Cells[row, 1].Value = t.Teminat;
                            ws.Cells[row, 2].Value = t.Bedel;
                            row++;
                        }

                        ws.Cells.AutoFitColumns();
                        excel.SaveAs(new FileInfo(excelPath));
                    }

                    SaveToExcel(poliçeNo, musteriNo, baslangicTarihi, bitisTarihi, tanzimTarihi, taksitler, ekBelgeNo, taksitTarihleri);

                    //MessageBox.Show($"Teminatlar Excel'e yazıldı: {excelPath}", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    string firmCode = "2"; // Varsayım: sabit değer veya GUI'den alınabilir
                    string companyName = cmbSirketSecimi.Text.Trim();
                    string productName = "BİLEŞİK ÜRÜN SİGORTA POLİÇESİ";
                    string curType = "TRY";

                    try
                    {
                        con.Open();


                        //Insert POLICYMASTER
                        OracleCommand cmd1 = new OracleCommand(@"
        INSERT INTO mercury2011.Z_PDFREADER_POLICYMASTER
        (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, CLIENT_NO, BEG_DATE, END_DATE, CONFIRM_DATE, CUR_TYPE)
        VALUES (:firm, :comp, :prod, :pol, :end, :cli, TO_DATE(:beg, 'DD/MM/YYYY'), TO_DATE(:fin, 'DD/MM/YYYY'), TO_DATE(:con, 'DD/MM/YYYY'), :cur)", con);

                        cmd1.Parameters.Add(":firm", firmCode);
                        cmd1.Parameters.Add(":comp", companyName);
                        cmd1.Parameters.Add(":prod", productName);
                        cmd1.Parameters.Add(":pol", poliçeNo);
                        cmd1.Parameters.Add(":end", ekBelgeNo);
                        cmd1.Parameters.Add(":cli", musteriNo);
                        cmd1.Parameters.Add(":beg", baslangicTarihi);
                        cmd1.Parameters.Add(":fin", bitisTarihi);
                        cmd1.Parameters.Add(":con", tanzimTarihi);
                        cmd1.Parameters.Add(":cur", curType);

                        cmd1.ExecuteNonQuery();

                        string[] taksitArr = taksitler.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < taksitArr.Length; i++)
                        {
                            //Tutarı temizle ve normalize et
                            string tutarStr = taksitArr[i]
                                .Replace("USD", "")
                                .Replace("TL", "")
                                .Replace("EUR", "");

                            if (decimal.TryParse(tutarStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tutar))
                            {
                                OracleCommand cmd2 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYINSTALLMENT
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, INSTALLMENT_ORDER, INSTALLMENT_DATE, INSTALLMENT_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :ord, :installdate, :amt)", con);

                                cmd2.Parameters.Add(":firm", firmCode);
                                cmd2.Parameters.Add(":comp", companyName);
                                cmd2.Parameters.Add(":prod", productName);
                                cmd2.Parameters.Add(":pol", poliçeNo);
                                cmd2.Parameters.Add(":end", ekBelgeNo); // Endors no = poliçeNo olabilir
                                cmd2.Parameters.Add(":ord", i + 1); // Taksit sırası
                                cmd2.Parameters.Add(":installdate", taksitTarihleri[i]); // Gerçek taksit tarihi varsa buraya koy
                                cmd2.Parameters.Add(":amt", tutar);

                                cmd2.ExecuteNonQuery();
                            }
                            else
                            {
                                MessageBox.Show($"Taksit tutarı geçersiz formatta: {taksitArr[i]}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }

                        //Insert POLICYCOVERAGE
                        foreach (var t in teminatlar)
                        {
                            string bedelStr = t.Bedel.Replace("TL", "").Replace("USD", "").Replace("EUR", "").Trim();
                            if (string.IsNullOrEmpty(bedelStr)) bedelStr = "0";

                            OracleCommand cmd3 = new OracleCommand(@"
            INSERT INTO mercury2011.Z_PDFREADER_POLICYCOVERAGE
            (FIRM_CODE, COMPANY_NAME, PRODUCT_NAME, POLICY_NO, ENDORS_NO, COVERAGE_NAME, COVERAGE_AMOUNT)
            VALUES (:firm, :comp, :prod, :pol, :end, :cov, :amt)", con);

                            cmd3.Parameters.Add(new OracleParameter(":firm", firmCode));
                            cmd3.Parameters.Add(new OracleParameter(":comp", companyName));
                            cmd3.Parameters.Add(new OracleParameter(":prod", productName));
                            cmd3.Parameters.Add(new OracleParameter(":pol", poliçeNo));
                            cmd3.Parameters.Add(new OracleParameter(":end", ekBelgeNo));
                            cmd3.Parameters.Add(new OracleParameter(":cov", t.Teminat));
                            cmd3.Parameters.Add(new OracleParameter(":amt", bedelStr));

                            cmd3.ExecuteNonQuery();
                        }

                        MessageBox.Show("SQL'e başarıyla kaydedildi.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("SQL Hatası: " + ex.Message);
                    }
                    finally
                    {
                        if (con.State == System.Data.ConnectionState.Open)
                            con.Close();
                    }

                }

                else
                {
                    MessageBox.Show("Yanlış Şirket PDF Seçimi", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtPDFContent.Clear();
                }


            }
        }



        private string GetUniqueFileName(string directory, string baseFileName)
        {
            string filePath = System.IO.Path.Combine(directory, baseFileName);
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(baseFileName);
            string extension = System.IO.Path.GetExtension(baseFileName);
            int counter = 1;

            while (File.Exists(filePath))
            {
                string newFileName = $"{fileNameWithoutExt}{counter}{extension}";
                filePath = System.IO.Path.Combine(directory, newFileName);
                counter++;
            }

            return filePath;
        }

        private void SaveToExcel(string poliçeNo, string musteriNo, string baslangicTarihi, string bitisTarihi, string tanzimTarihi, string taksitler, string ekBelgeNo, List<DateTime> taksitTarihleri)
        {
            string saveDirectory = @"C:\EXCEL";

            if (!Directory.Exists(saveDirectory))
            {
                Directory.CreateDirectory(saveDirectory);
            }

            string selectedCompany = cmbSirketSecimi.Text.Trim().ToUpper();
            string baseFileName = "";

            if (selectedCompany.Contains("AXA"))
            {
                baseFileName = "axa_pdf.xlsx";
            }
            else if (selectedCompany.Contains("SOMPO"))
            {
                baseFileName = "sompo_pdf.xlsx";
            }
            else if (selectedCompany.Contains("TURKIYE"))
            {
                baseFileName = "turkiye_pdf.xlsx";
            }
            else
            {
                baseFileName = "default_pdf.xlsx";
            }

            // Benzersiz dosya yolu alın
            string excelPath = GetUniqueFileName(saveDirectory, baseFileName);

            using (var excel = new OfficeOpenXml.ExcelPackage())
            {
                var ws = excel.Workbook.Worksheets.Add("Poliçe Bilgileri");

                ws.Cells["A1"].Value = "Poliçe No";
                ws.Cells["B1"].Value = "Müşteri No";
                ws.Cells["C1"].Value = "Başlangıç Tarihi";
                ws.Cells["D1"].Value = "Bitiş Tarihi";
                ws.Cells["E1"].Value = "Tanzim Tarihi";
                ws.Cells["F1"].Value = "Taksitler";
                ws.Cells["G1"].Value = "Taksit Tarihleri";
                ws.Cells["H1"].Value = "Teyzil";

                ws.Cells["A2"].Value = poliçeNo;
                ws.Cells["B2"].Value = musteriNo;
                ws.Cells["C2"].Value = baslangicTarihi;
                ws.Cells["D2"].Value = bitisTarihi;
                ws.Cells["E2"].Value = tanzimTarihi;
                ws.Cells["H2"].Value = ekBelgeNo;

                if (!string.IsNullOrEmpty(taksitler))
                {
                    var taksitList = taksitler.Split(new[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    if (taksitList.Length == 0)
                        MessageBox.Show("BOŞ");

                    else
                    {
                        int startRow = 2; // Başlangıç satırı
                        int taksitColumn = 6; // F sütunu (6. sütun)

                        for (int i = 0; i < taksitList.Length; i++)
                        {
                            int row = startRow + i;
                            ws.Cells[row, taksitColumn].Value = taksitList[i].Trim();
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Taksit Yok");
                }


                if (taksitTarihleri != null && taksitTarihleri.Count > 0)
                {
                    int startRow = 2;
                    int tarihColumn = 7;

                    for (int i = 0; i < taksitTarihleri.Count; i++)
                    {
                        int row = startRow + i;
                        ws.Cells[row, tarihColumn].Value = taksitTarihleri[i].ToString("dd/MM/yyyy");
                    }
                }

                ws.Cells.AutoFitColumns();

                excel.SaveAs(new FileInfo(excelPath));
                //MessageBox.Show($"Poliçe Excel'e yazıldı: {excelPath}", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        //Yardımcı fonksiyon
        private string GetNextValue(string[] lines, int index)
        {
            if (index + 1 < lines.Length)
            {
                var match = Regex.Match(lines[index + 1], @"^\s*:\s*(.+)$");
                if (match.Success)
                    return match.Groups[1].Value.Trim();
            }
            return "";

        }

        private string ExtractValue(string text, string fieldName)
        {
            try
            {
                string pattern;
                if (fieldName == "Poliçe No")
                {
                    pattern = @"Poliçe No:\s*(\d+)";
                }
                else if (fieldName == "Müşteri No")
                {
                    pattern = @"Müşteri No:\s*(\d+)";
                }
                else if (fieldName == "Başlangıç Tarihi")
                {
                    pattern = @"Başlangıç Tarihi:\s*([0-9]{2}/[0-9]{2}/[0-9]{4}(?: [0-9]{2}:[0-9]{2})?)";
                }
                else if (fieldName == "Bitiş Tarihi")
                {
                    pattern = @"Bitiş Tarihi:\s*([0-9]{2}/[0-9]{2}/[0-9]{4}(?: [0-9]{2}:[0-9]{2})?)";
                }
                else if (fieldName == "Tanzim Tarihi")
                {
                    pattern = @"Tanzim Tarihi\s*:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})";
                }
                else
                {
                    pattern = $@"{fieldName}\s*:\s*([^\s]+)";
                }

                var match = Regex.Match(text, pattern);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata oluştu: {ex.Message}");
            }
            return "";
        }

        private void cmbBelgeTuru_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbSirketSecimi.Items.Clear();

            if (cmbBelgeTuru.Text == "Poliçe")
            {
                con.Open();
                OracleCommand cmd = new OracleCommand("SELECT DESCRIPTION FROM MERCURY2011.DEFINITIONDB WHERE TYPE_CODE = 'SGS' ORDER BY DESCRIPTION", con);
                OracleDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    cmbSirketSecimi.Items.Add(rd.GetValue(0).ToString());
                }
                con.Close();
                cmbSirketSecimi.Enabled = true;

            }
            else if (cmbBelgeTuru.Text == "Credit Note")
            {
                cmbSirketSecimi.Items.Add("test");
                cmbSirketSecimi.Enabled = true;
            }
            else if (cmbBelgeTuru.Text == "Vergi Levhası")
            {
                cmbSirketSecimi.Items.Add("test2");
                cmbSirketSecimi.Enabled = true;
            }



        }


    }



}
