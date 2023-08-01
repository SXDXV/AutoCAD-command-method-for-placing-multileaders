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
using System.IO;

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
        Dictionary<int, string> actualBlockOrderDict = new Dictionary<int, string>();
        Dictionary<string, int> actualBlockOrderDictDublicates = new Dictionary<string, int>();
        Dictionary<int, string> selectedBlockOrderDict = new Dictionary<int, string>();
        List<Block> blocks = new List<Block>();

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


            Creating(doc, ed);
            try
            {
                
            }
            catch (System.Exception undefinedError)
            {
                ed.WriteMessage("Ошибка неопределенного поведения программы. " +
                    "Обратитесь в отдел разработки расширения.\nLOGS: " + undefinedError.ToString());
            }
        }

        private void Creating(Document doc, Editor ed)
        {
            actualBlockOrderDictDublicates.Clear();
            actualBlockOrderDict.Clear();
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
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // Переменная для хранения порядкового номера
                        int visualiseNumber = 1;

                        foreach (ObjectId objectId in selectionSet.GetObjectIds())
                        {
                            int orderNumber = visualiseNumber;
                            BlockReference blockRef = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                            if (blockRef != null)
                            {
                                string blockID = blockRef.XData.AsArray()[2].Value.ToString();
                                int foundKey = 0;

                                // Проверяем, есть ли блок в словаре
                                if (selectedBlockOrderDict.ContainsValue(blockID))
                                {
                                    foreach (var kvp in selectedBlockOrderDict)
                                    {
                                        if (kvp.Value == blockID)
                                        {
                                            foundKey = kvp.Key;
                                        }
                                    }
                                    // Если блок уже есть в словаре, берем его порядковый номер
                                    orderNumber = foundKey;
                                }
                                else
                                {
                                    if (selectedBlockOrderDict.Any())
                                    {
                                        orderNumber = selectedBlockOrderDict.Keys.Max() + 1;
                                        // Если блока еще нет в словаре, добавляем его с текущим порядковым номером
                                        selectedBlockOrderDict.Add(orderNumber, blockID);
                                    }
                                    else
                                    {
                                        // Если блока еще нет в словаре, добавляем его с текущим порядковым номером
                                        selectedBlockOrderDict.Add(orderNumber, blockID);
                                    }

                                }

                                // Создание MLeader (мультивыноски)
                                createMleader(orderNumber, blockRef, db, modelSpace, tr);

                                visualiseNumber++;
                            }
                        }

                        foreach (ObjectId objectId in modelSpace)
                        {
                            if (objectId.ObjectClass == RXObject.GetClass(typeof(MLeader)))
                            {
                                MLeader leader = tr.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                if (leader != null)
                                {
                                    // Обработка мультивыносок
                                    string mTextContents = leader.MText.Contents;

                                    // Проверяем, есть ли значение mTextContents в словаре blockOrderDict
                                    if (selectedBlockOrderDict.ContainsKey(Convert.ToInt32(mTextContents)))
                                    {
                                        try 
                                        {
                                            actualBlockOrderDict.Add(Convert.ToInt32(mTextContents), selectedBlockOrderDict[Convert.ToInt32(mTextContents)]);
                                            actualBlockOrderDictDublicates.Add(selectedBlockOrderDict[Convert.ToInt32(mTextContents)], 1);
                                        }
                                        catch
                                        {
                                            actualBlockOrderDictDublicates[selectedBlockOrderDict[Convert.ToInt32(mTextContents)]] =
                                                actualBlockOrderDictDublicates[selectedBlockOrderDict[Convert.ToInt32(mTextContents)]] + 1;
                                        }
                                    }
                                }
                            }
                        }

                        CreateTableWithFields();

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


            //foreach (var kvp in actualBlockOrderDict)
            //{
            //    selectedBlockOrderDict[kvp.Key] = kvp.Value;
            //}
            //oldBlockOrderDict.addRange(actualBlockOrderDict);
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

        private void createMleader(int orderNumber, BlockReference blockRef, Database db, BlockTableRecord modelSpace, Transaction tr)
        {
            MLeader leader = new MLeader();
            leader.SetDatabaseDefaults();
            leader.ContentType = ContentType.MTextContent;

            MText mText = new MText();
            mText.SetDatabaseDefaults();
            mText.TextHeight = 20.0;
            mText.Contents = orderNumber.ToString(); // Нумеруем выноски

            Point3d mTextPosition = new Point3d(blockRef.Position.X + 15, blockRef.Position.Y + GetBlockHeight(blockRef), 0);
            mText.Location = mTextPosition;
            leader.MText = mText;
            leader.MLeaderStyle = CustomMLeaderStyle(db);

            Point3d firstPoint = new Point3d(blockRef.Position.X, blockRef.Position.Y, 0);
            int idx = leader.AddLeaderLine(firstPoint);
            leader.AddFirstVertex(idx, firstPoint);

            modelSpace.AppendEntity(leader);
            tr.AddNewlyCreatedDBObject(leader, true);
        }

        public void ParseDataToArray()
        {
            string appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string databasePath = "PartsDataBase.sqlite";
            string connectionString = $"Data Source={appPath}\\{databasePath};Version=3;";

            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                foreach (KeyValuePair<int, string> kvp in selectedBlockOrderDict)
                {
                    int number = kvp.Key;
                    string id = kvp.Value;

                    Block block = GetBlockInfo(connection, id);

                    if (block != null)
                    {
                        blocks.Add(block);
                    }
                }
            }
        }

        private static Block GetBlockInfo(SQLiteConnection connection, string id)
        {
            string selectQuery = "SELECT * FROM 'Parts' WHERE ID = " + '"' + id + '"' + ";";

            using (SQLiteCommand command = new SQLiteCommand(selectQuery, connection))
            {

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    Block block = null;
                    while (reader.Read())
                    {
                        block = new Block(reader.GetValue(0).ToString(), Convert.ToDouble(reader.GetValue(1)), Convert.ToInt32(reader.GetValue(2)), reader.GetValue(3).ToString());
                    }
                    return block;
                }
            }
        }

        public void CreateTableWithFields()
        {
            ParseDataToArray();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Задаем параметры таблицы
            int numRows = actualBlockOrderDict.Count + 2; // Количество строк
            int numCols = 4; // Количество столбцов
            double startX = -20500; // Начальная координата X
            double startY = 26430; // Начальная координата Y

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // Создаем таблицу
                Table table = new Table();
                table.SetDatabaseDefaults();
                table.Position = new Point3d(startX, startY, 0.0);
                table.NumRows = numRows;
                table.NumColumns = numCols;
                table.TableStyle = db.Tablestyle;
                table.GenerateLayout();

                // Добавляем таблицу в пространство модели
                modelSpace.AppendEntity(table);
                tr.AddNewlyCreatedDBObject(table, true);

                int[] customColumnWidths = new int[] { 200, 1000, 200, 200 };

                // Заполняем таблицу данными
                for (int row = 0; row < numRows; row++)
                {
                    table.Rows[row].Height = 50;
                    for (int col = 0; col < numCols; col++)
                    {
                        table.Columns[col].Width = customColumnWidths[col];

                        Cell cell = table.Cells[row, col];
                        cell.TextHeight = 20;

                        if (row == 1)
                        {
                            switch (col)
                            {
                                case 0:
                                    cell.TextString = "№";
                                    break;
                                case 1:
                                    cell.TextString = "Наименование";
                                    break;
                                case 2:
                                    cell.TextString = "Количество";
                                    break;
                                case 3:
                                    cell.TextString = "Масса";
                                    break;
                            }
                        }
                        else if (row > 1)
                        {
                            int number = row - 1;
                            switch (col)
                            {
                                case 0:
                                    cell.TextString = number.ToString(); ;
                                    break;
                                case 1:
                                    cell.TextString = blocks[number-1].Fullname
                                        .Replace("<$D$>", blocks[number - 1].Diameter.ToString())
                                        .Replace("<$D1$>", blocks[number - 1].Diameter.ToString());
                                    break;
                                case 2:
                                    try
                                    {
                                        cell.TextString = actualBlockOrderDictDublicates[blocks[number - 1].Id].ToString();
                                    }
                                    catch
                                    {
                                        cell.TextString = "1";
                                    }
                                    break;
                                case 3:
                                    cell.TextString = blocks[number - 1].Weight.ToString();
                                    break;
                            }
                        }
                    }
                }

                tr.Commit();
            }

            ed.Regen(); // Обновляем отображение
        }
    }
}
