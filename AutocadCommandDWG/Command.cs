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

// Автор проекта - Шевцов Дмитрий (SXDXV)

// Проект, созданн в целях демонстрации навыков работы с ЯП C# и умении работать
// с документацией незнакомой технологии. Конкретно в рамках проекта осуществяется подключение
// к AutoCAD посредством взаимодействия с API продукта.

namespace AutocadCommandDWG
{
    /// <summary>
    /// Главный класс команды.
    /// На данном этапе проверяется проверка интегрируемости продукта в систему AutoCAD.
    /// Главная задача на данном этапе - првоерить любой минимальный исход программы.
    /// </summary>
    public class Command
    {
        /// <summary>
        /// Непосредственно метод взаимодействия с командной строкой, который на данном этапе 
        /// разработки программного обеспечения реализует фунционал обратного отклика пользователю о
        /// выбранных / не выбранных блоках, а также отлавливает непредвиденные ситуации.
        /// </summary>
        [CommandMethod("SelectBlocks")]
        public void SelectBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult selectionResult = ed.GetSelection();
                if (selectionResult.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionResult.Value;
                    if (selectionSet.Count > 0)
                    {
                        ed.WriteMessage("Были выбраны вхождения блоков.");

                        Database db = doc.Database;
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                            if (selectionResult.Status == PromptStatus.OK)
                            {
                                // Переменная для хранения порядкового номера
                                int orderNumber = 1;

                                foreach (ObjectId objectId in selectionSet.GetObjectIds())
                                {
                                    BlockReference blockRef = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                                    if (blockRef != null)
                                    {
                                        MLeader leader = new MLeader();
                                        leader.SetDatabaseDefaults();
                                        leader.ContentType = ContentType.MTextContent;

                                        MText mText = new MText();
                                        mText.SetDatabaseDefaults();
                                        mText.Width = 100;
                                        mText.Height = 50;
                                        mText.Contents = orderNumber.ToString(); // Нумеруем выноски
                                        mText.Location = blockRef.Position.Add(new Vector3d(100, 100, GetBlockHeight(blockRef, tr))); // Поднимаем выноску над блоком

                                        leader.MText = mText;

                                        int idx = leader.AddLeaderLine(blockRef.Position.Add(new Vector3d(100, 100, 0))); // Присоединяем линию к блоку
                                        leader.AddFirstVertex(idx, blockRef.Position);

                                        modelSpace.AppendEntity(leader);
                                        tr.AddNewlyCreatedDBObject(leader, true);

                                        orderNumber++;
                                    }
                                }
                            }

                            tr.Commit();
                        }

                        doc.Editor.Regen();
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
            catch (System.Exception undefinedError)
            {
                ed.WriteMessage("Ошибка неопределенного поведения программы. " +
                    "Обратитесь в отдел разработки расширения.\nLOGS: " + undefinedError.ToString());
            }
        }

        private double GetBlockHeight(BlockReference blockRef, Transaction tr)
        {
            // Получаем имя блока из ссылки на блок
            string blockName = blockRef.Name;

            // Получаем определение блока из базы данных чертежа
            BlockTableRecord blockDef = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

            // Проверяем, есть ли атрибут "HEIGHT" в определении блока
            if (blockDef != null && blockDef.HasAttributeDefinitions)
            {
                foreach (ObjectId attId in blockDef)
                {
                    AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null && attDef.Constant && attDef.Tag.ToUpper() == "HEIGHT")
                    {
                        // Находим атрибут "HEIGHT" и возвращаем его значение
                        using (AttributeReference attRef = new AttributeReference())
                        {
                            attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                            return attRef.Height;
                        }
                    }
                }
            }

            // Возвращаем значение по умолчанию, если атрибут "HEIGHT" не найден
            return 0.0;
        }
    }
}
