using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using DocumentFormat.OpenXml;

namespace OpenXml
{
    class Program
    {
        static void Main(string[] args)
        {

            var dictionary = new Dictionary<string, string>();
            dictionary.Add("FullName", "Patrick Wolfs");
            dictionary.Add("Address", "Hasseltsestraat 40");
            dictionary.Add("PostalCode", "3540");
            dictionary.Add("City", "Herk-de-Stad");
            dictionary.Add("Country", "België");

            var filePath = @"D:\Users\Patrick\Downloads\ModemSwap.docx";
            using (var docPackage = WordprocessingDocument.Open(filePath, true))
            {
                string fieldList = string.Empty;

                var document = docPackage.MainDocumentPart.Document;
                IEnumerable<FieldChar> fields = document.Descendants<FieldChar>();
                if (fields == null) return; //No field codes in the document

                // bool fldStart = false;
                FieldChar fldCharStart = null;
                FieldChar fldCharEnd = null;
                FieldChar fldCharSep = null;
                FieldCode fldCode = null;
                string fldContent = String.Empty;
                foreach (FieldChar fldChar in fields)
                {
                    string fldCharPart = fldChar.FieldCharType.ToString();
                    switch (fldCharPart)
                    {
                        case "begin": //start of the field
                            fldCharStart = fldChar;
                            //get the field code, which will be an instrText element
                            // either as sibling or as a child of the parent sibling
                            fldCode = fldCharStart.Parent.Descendants<FieldCode>().FirstOrDefault();
                            if (fldCode == null) //complex field
                            {
                                fldCode =
                                    fldCharStart.Parent.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>()
                                        .Descendants<FieldCode>()
                                        .FirstOrDefault();
                            }
                            if (fldCode != null && fldCode.InnerText.Contains("MERGEFIELD"))
                            {
                                var key = dictionary.Keys.FirstOrDefault(k => fldCode.InnerText.Contains(k));
                                if (!string.IsNullOrEmpty(key))
                                {
                                    fldContent = dictionary[key];
                                }
                            }
                            break;
                        case "end": // end of the field
                            fldCharEnd = fldChar;
                            break;
                        case "separate": //complex field with text result
                            //we want to put the database content in this text run
                            //yet still remove the field code
                            //If there's no "separate" field char for the current field,
                            //we need to insert it somewhere else
                            fldCharSep = fldChar;
                            break;
                        default:
                            break;
                    }
                    if ((fldCharStart != null) && (fldCharEnd != null)) //start and end field codes have been found
                    {
                        if (fldCharSep != null)
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Text elemText =
                                (DocumentFormat.OpenXml.Wordprocessing.Text)
                                fldCharSep?.Parent?.NextSibling()?
                                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                                    .FirstOrDefault();
                            if (elemText != null)
                            {
                                elemText.Text = fldContent;
                            }
                            //Delete all the field chars with their runs
                            DeleteFieldChar(fldCharStart);
                            DeleteFieldChar(fldCharEnd);
                            DeleteFieldChar(fldCharSep);
                            if (fldCode.Parent != null)
                                fldCode.Remove();
                        }
                        else
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Text elemText = new DocumentFormat.OpenXml.Wordprocessing.Text(fldContent);
                            fldCode.Parent.Append(elemText);
                            fldCode.Remove();
                            //Delete all the field chars with their runs
                            DeleteFieldChar(fldCharStart);
                            DeleteFieldChar(fldCharEnd);
                            DeleteFieldChar(fldCharSep);
                        }
                        fldCharStart = null;
                        fldCharEnd = null;
                        fldCharSep = null;
                        fldCode = null;
                        fldContent = string.Empty;
                    }
                }
                document.Save();
            }
        }

        private static void DeleteFieldChar(OpenXmlElement fldCharStart)
        {
            Run fldRun = (Run) fldCharStart.Parent;
            if (fldRun == null) return;
            fldRun.RemoveAllChildren();
            fldRun.Remove();
        }
    }
}
