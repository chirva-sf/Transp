using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Transp
{
    class ClWord
    {
        public static Word.Application hWord;
        public static Word.Document hDoc;
        public static bool NeedActivate = true;

        public ClWord()
        {
            hWord = (Word.Application)StartOrGetWord();
        }

        private static object StartOrGetWord()
        {
            object hRes = null;
            try
            {
                hRes = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                if (hRes == null) hRes = new Word.ApplicationClass();
            }
            return hRes;
        }

        public static bool Open(object TemplateFile)
        {
            if (hWord == null || TemplateFile == null) return false;
            return Open(TemplateFile, false, Word.WdNewDocumentType.wdNewBlankDocument);
        }

        public static bool OpenForEdit(object TemplateFile)
        {
            if (hWord == null || TemplateFile == null) return false;
            object NO = System.Reflection.Missing.Value;
            hWord.Documents.Open(ref TemplateFile, ref NO, ref NO, ref NO, ref NO,
                ref NO, ref NO, ref NO, ref NO, ref NO, ref NO, ref NO, ref NO, ref NO,
                ref NO);
            return true;
        }

        private static bool Open(object TemplateFile, object AsTemplate, object TypeOpen)
        {
            if (hWord == null || TemplateFile == null) return false;
            object DocVisible = true;
            while (true)
            {
                try
                {
                    hDoc = hWord.Documents.Add(ref TemplateFile, ref AsTemplate, ref TypeOpen, ref DocVisible);
                    break;
                }
                catch (Exception E)
                {
                    if (E.Message == "Сервер RPC недоступен.")
                    {
                        hWord = null;
                        hWord = (Word.Application)StartOrGetWord();
                        continue;
                    }
                    if (E.Message.IndexOf("диалоговое окно ''Найти'' или ''Заменить'' открыто") > 0)
                    {
                        return false;
                    }
                    if (E.Message.StartsWith("Неверно указаны путь или имя документа"))
                        return false;
                    if (E.Message.StartsWith("Недостаточно памяти. Немедленно сохраните документ."))
                    {
                        return false;
                    }
                    return false;
                }
            }
            return (hDoc != null);
        }

        public static void Complete()
        {
            try
            {
                hDoc.PrintPreview();
                hDoc.ClosePrintPreview();
            }
            catch
            {
            }
            try
            {
                hDoc.Fields.Update();
            }
            catch
            {
            }
            if (NeedActivate)
                CompleteForEdit();
        }

        public static void CompleteForEdit()
        {
            try
            {
                hWord.Visible = true;
                hWord.Activate();
                for (int i = hWord.Windows.Count; i >= 1; i--)
                {
                    object nC = i;
                    hWord.Windows.Item(ref nC).Activate();
                }
            }
            catch
            {
            }
        }

        public static void SetVar(object VarName, string VarValue)
        {
            try
            {
                hDoc.Variables.Item(ref VarName).Value = VarValue == "" ? " " : VarValue;
            }
            catch
            {

            }
        }

        public static void InsertLine(int TabNom, int BeforeLineNum)
        {
            object beforeRow = hDoc.Tables.Item(TabNom).Cell(BeforeLineNum, 1);
            hDoc.Tables.Item(TabNom).Rows.Add(ref beforeRow);
        }

        public static void AddRow(int TabNom)
        {
            hDoc.Tables.Item(TabNom).Rows.Add();
        }

        public static void SetCellValue(int TabNom, int x, int y, string sVal)
        {
            hDoc.Tables.Item(TabNom).Cell(y, x).Range.Text = sVal;
        }
    }
}
