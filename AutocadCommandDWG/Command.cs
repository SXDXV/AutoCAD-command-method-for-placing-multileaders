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
using Autodesk.AutoCAD.Colors;

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
        Dictionary<string, int> blockOrderDict = new Dictionary<string, int>();

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
                // Создаем фильтр для выбора только блоков
                PromptSelectionOptions opts = new PromptSelectionOptions();
                SelectionFilter filter = new SelectionFilter(
                    new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") }
                );
                PromptSelectionResult selectionResult = ed.GetSelection(opts, filter);

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

                            // Переменная для хранения порядкового номера
                            int actualNumber = 1;

                            foreach (ObjectId objectId in selectionSet.GetObjectIds())
                            {
                                int orderNumber = actualNumber;
                                BlockReference blockRef = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                                if (blockRef != null)
                                {
                                    // Проверяем, есть ли блок в словаре
                                    if (blockOrderDict.ContainsKey(blockRef.Name))
                                    {
                                        // Если блок уже есть в словаре, берем его порядковый номер
                                        orderNumber = blockOrderDict[blockRef.Name];
                                    }
                                    else
                                    {
                                        if (blockOrderDict.Any())
                                        {
                                            orderNumber = blockOrderDict.Values.Max()+1;
                                            // Если блока еще нет в словаре, добавляем его с текущим порядковым номером
                                            blockOrderDict.Add(blockRef.Name, orderNumber);
                                        }
                                        else
                                        {
                                            // Если блока еще нет в словаре, добавляем его с текущим порядковым номером
                                            blockOrderDict.Add(blockRef.Name, orderNumber);
                                        }
                                        
                                    }

                                    double blockHeight = GetBlockHeight(blockRef);
                                    double blockWidth = GetBlockWidth(blockRef);

                                    MLeader leader = new MLeader();
                                    leader.SetDatabaseDefaults();
                                    leader.ContentType = ContentType.MTextContent;

                                    MText mText = new MText();
                                    mText.SetDatabaseDefaults();
                                    mText.TextHeight = 20.0;
                                    mText.Contents = orderNumber.ToString(); // Нумеруем выноски

                                    Point3d mTextPosition = new Point3d(blockRef.Position.X + 15, blockRef.Position.Y + blockHeight, 0);
                                    mText.Location = mTextPosition;
                                    leader.MText = mText;
                                    leader.MLeaderStyle = CustomMLeaderStyle(db);

                                    Point3d firstPoint = new Point3d(blockRef.Position.X, blockRef.Position.Y, 0);
                                    int idx = leader.AddLeaderLine(firstPoint);
                                    leader.AddFirstVertex(idx, firstPoint);

                                    modelSpace.AppendEntity(leader);
                                    tr.AddNewlyCreatedDBObject(leader, true);

                                    actualNumber++;
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

        private double GetBlockHeight(BlockReference blockRef)
        {
            // Получаем ограничивающий прямоугольник блока
            Extents3d extents = blockRef.GeometricExtents;

            // Вычисляем высоту блока по оси Y
            double height = extents.MaxPoint.Y - extents.MinPoint.Y;

            return height;
        }

        private double GetBlockWidth(BlockReference blockRef)
        {
            // Получаем ограничивающий прямоугольник блока
            Extents3d extents = blockRef.GeometricExtents;

            // Вычисляем ширину блока по оси X
            double width = extents.MaxPoint.X - extents.MinPoint.X;

            return width;
        }

        private ObjectId CustomMLeaderStyle(Database db)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeaderStyle mLeaderStyle = new MLeaderStyle();

                // Задаем параметры стиля мультивыноски
                mLeaderStyle.ArrowSymbolId = ObjectId.Null; // Идентификатор блока для стрелки (может быть изменен на нужный блок)
                mLeaderStyle.ContentType = ContentType.MTextContent;
                mLeaderStyle.TextHeight = 3; // Высота текста для цифры мультивыноски
                mLeaderStyle.EnableLanding = true; // Включаем отображение платформы
                mLeaderStyle.LandingGap = 10; // Расстояние между текстом и платформой

                // Добавляем стиль мультивыноски в базу данных чертежа
                DBDictionary mlStyleDict = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForWrite) as DBDictionary;
                ObjectId styleId = mlStyleDict.SetAt("CustomMLeaderStyle", mLeaderStyle);

                tr.AddNewlyCreatedDBObject(mLeaderStyle, true);
                tr.Commit();

                return styleId;
            }
        }
    }
}
