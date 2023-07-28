using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Data;
using System.Data.SQLite;
using System.Xml.Linq;

namespace AutocadCommandDWG
{
    public class Command
    {
        [CommandMethod("SelectBlocks")]
        public void SelectBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptSelectionResult selectionResult = ed.GetSelection();
            if (selectionResult.Status == PromptStatus.OK)
            {
                SelectionSet selectionSet = selectionResult.Value;
                if (selectionSet.Count > 0)
                {
                    ed.WriteMessage("Были выбраны вхождения блоков.");
                }
                else
                {
                    ed.WriteMessage("Нет выбранных вхождений блоков.");
                }
            }
            else if (selectionResult.Status == PromptStatus.Cancel)
            {
                ed.WriteMessage("Выбор отменен.");
            }
            else if (selectionResult.Status == PromptStatus.None)
            {
                ed.WriteMessage("Ничего не выбрано.");
            }
            else if (selectionResult.Status == PromptStatus.Error)
            {
                ed.WriteMessage("Ошибка при выборе.");
            }
        }
    }
}
