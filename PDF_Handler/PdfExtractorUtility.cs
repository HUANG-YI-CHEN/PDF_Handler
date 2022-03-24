using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PDF_Handler
{
    class PdfExtractorUtility
    {
        /// <summary> 
        /// 從已有PDF檔案中拷貝指定的頁碼範圍到一個新的pdf檔案中 使用pdfCopyProvider.AddPage()方法
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        public void SplitPDF(string sourcePdfPath, string outputPdfPath, int startPage, int endPage)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;
            try
            {
                reader = new PdfReader(sourcePdfPath);
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage)); pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
                sourceDocument.Open();
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i); pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex) { throw ex; }
        }


        /// <summary> 
        /// 將PDF檔案分割成單頁
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        public void Split2SinglePage(string sourcePdfPath)
        {
            PdfReader reader = null;
            try
            {
                string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(sourcePdfPath);
                string outputPdfFolder = System.IO.Path.GetDirectoryName(sourcePdfPath);
                reader = new PdfReader(sourcePdfPath);

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfCopy pdfCopyProvider = null;
                    PdfImportedPage importedPage = null;
                    Document sourceDocument = null;
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + "_" + i + ".pdf";
                    sourceDocument = new Document(reader.GetPageSizeWithRotation(i));
                    pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
                    sourceDocument.Open();
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                    sourceDocument.Close();
                }

                reader.Close();
            }
            catch (Exception ex) { throw ex; }
        }


        /// <summary> 
        /// 將PDF檔案平均分割成多個檔案，無法分盡，剩餘頁數就加到最後一個檔案
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        /// <param name="count">需要生成的檔案數量</param>
        public void Split2AveragePage(string sourcePdfPath, int count)
        {
            PdfReader reader = null;
            try
            {
                string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(sourcePdfPath);
                string outputPdfFolder = System.IO.Path.GetDirectoryName(sourcePdfPath);
                reader = new PdfReader(sourcePdfPath);
                // int page = (reader.NumberOfPages / count);
                // 計算每個檔案的頁數，總是捨去小數
                int page = (int)Math.Floor((double)(reader.NumberOfPages) / (double)(count));
                int startPage = 1;
                int endPage = 1;

                LogUtil.WriteLog("每個檔案的頁數：" + page.ToString());

                for (int i = 1; i <= count; i++)
                {
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + "_" + i + ".pdf"; ;

                    if (i == 1)
                    {
                        startPage = 1;
                        endPage = page;

                    }
                    else
                    {

                        startPage = endPage + 1;
                        endPage = startPage + page - 1;
                    }

                    if (startPage > reader.NumberOfPages)
                        break;

                    if (endPage > reader.NumberOfPages)
                        endPage = reader.NumberOfPages;

                    if (i == count)
                        endPage = reader.NumberOfPages;

                    LogUtil.WriteLog(outputPdfPath + " > " + startPage.ToString() + "-" + endPage.ToString());
                    SplitPDF(sourcePdfPath, outputPdfPath, startPage, endPage);

                }

                reader.Close();
            }
            catch (Exception ex) { throw ex; }
        }



        /// <summary> 
        /// 將PDF檔案按檔案固定頁數割成多個檔案
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        /// <param name="page">每個檔案頁數</param>
        public void Split2Page(string sourcePdfPath, int page)
        {
            PdfReader reader = null;
            try
            {
                string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(sourcePdfPath);
                string outputPdfFolder = System.IO.Path.GetDirectoryName(sourcePdfPath);
                reader = new PdfReader(sourcePdfPath);
                // int page = (reader.NumberOfPages / count);
                // 計算按固定頁數生成檔案的數量 只要有小數都加1
                int count = (int)Math.Ceiling((double)(reader.NumberOfPages) / (double)(page));
                int startPage = 1;
                int endPage = 1;

                LogUtil.WriteLog("檔案數量：" + count.ToString());

                for (int i = 1; i <= count; i++)
                {
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + "_" + i + ".pdf"; ;

                    if (i == 1)
                    {
                        startPage = 1;
                        endPage = page;

                    }
                    else
                    {

                        startPage = endPage + 1;
                        endPage = endPage + page;
                    }

                    if (startPage > reader.NumberOfPages)
                        break;

                    if (endPage > reader.NumberOfPages)
                        endPage = reader.NumberOfPages;

                    if (i == count)
                        endPage = reader.NumberOfPages;

                    LogUtil.WriteLog(outputPdfPath + " > " + startPage.ToString() + "-" + endPage.ToString());
                    SplitPDF(sourcePdfPath, outputPdfPath, startPage, endPage);

                }

                reader.Close();
            }
            catch (Exception ex) { throw ex; }
        }


        /// <summary> 
        /// 從已有PDF檔案中拷貝指定的頁碼範圍到一個新的pdf檔案中 使用pdfCopyProvider.AddPage()方法
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        /// <param name="custpages">自定義的頁數範圍</param>
        public void SplitPDFCustPage(string sourcePdfPath, string custpages)
        {
            //  string[] strArray = custpages.Trim().Split(",");
            string[] strArray = custpages.Trim().Split(new Char[] { ',' });
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(sourcePdfPath);
            string outputPdfFolder = System.IO.Path.GetDirectoryName(sourcePdfPath);
            int startPage;
            int endPage;

            for (int i = 0; i < strArray.Length; i++)
            {

                LogUtil.WriteLog("自定義頁面範圍：" + strArray[i]);

                // 橫槓-相連的頁碼，抽取連續的範圍內的頁碼生成到一個檔案
                if (strArray[i].Contains("-"))
                {
                    //  string[] array = strArray[i].Split("-");
                    string[] array = strArray[i].Split(new Char[] { '-' });
                    startPage = int.Parse(array[0]);
                    endPage = int.Parse(array[1]);
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + " " + startPage + "-" + endPage + ".pdf";
                    LogUtil.WriteLog(outputPdfPath);
                    SplitPDF(sourcePdfPath, outputPdfPath, startPage, endPage);

                }
                // and &相連的頁碼，抽取指定頁碼生成到一個檔案
                else if (strArray[i].Contains("&"))
                {
                    //   int[] intArray = Array.ConvertAll(strArray[i].Split("&"), int.Parse);
                    int[] intArray = Array.ConvertAll(strArray[i].Split(new Char[] { '&' }), int.Parse);
                    string pages = string.Join("&", intArray);
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + " " + pages + ".pdf";
                    LogUtil.WriteLog(outputPdfPath);
                    SplitPDF2ExtractPages(sourcePdfPath, outputPdfPath, intArray);

                }
                else
                {
                    startPage = int.Parse(strArray[i]);
                    endPage = int.Parse(strArray[i]);
                    string outputPdfPath = outputPdfFolder + "\\" + fileNameWithoutExtension + " " + strArray[i] + ".pdf"; ;
                    LogUtil.WriteLog(outputPdfPath);
                    SplitPDF(sourcePdfPath, outputPdfPath, startPage, endPage);
                }
            }
        }

        /// <summary> 
        /// 將已有pdf檔案中 不連續 的頁拷貝至新的pdf檔案中。其中需要拷貝的頁碼存於陣列 int[] extractThesePages中
        /// </summary>
        /// <param name="sourcePdfPath">檔案路徑+檔名</param>
        /// <param name="extractThesePages">頁碼集合</param>
        /// <param name="outputPdfPath">檔案路徑+檔名</param>
        public void SplitPDF2ExtractPages(string sourcePdfPath, string outputPdfPath, int[] extractThesePages)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                reader = new PdfReader(sourcePdfPath);
                sourceDocument = new Document(reader.GetPageSizeWithRotation(extractThesePages[0]));
                pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
                sourceDocument.Open();
                foreach (int pageNumber in extractThesePages)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber); pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
