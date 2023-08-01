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
        Dictionary<int, string> actualBlockOrderDict = new Dictionary<int, string>();
        Dictionary<int, string> selectedBlockOrderDict = new Dictionary<int, string>();

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
                //Creating(doc, ed);
            }
            catch (System.Exception undefinedError)
            {
                ed.WriteMessage("Ошибка неопределенного поведения программы. " +
                    "Обратитесь в отдел разработки расширения.\nLOGS: " + undefinedError.ToString());
            }
        }

        private void Creating(Document doc, Editor ed)
        {
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

                                CreateTableWithFields();
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
                                        actualBlockOrderDict.Add(Convert.ToInt32(mTextContents), selectedBlockOrderDict[Convert.ToInt32(mTextContents)]);
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

        public void parseDataToArray()
        {
            string databasePath = "PartsDataBase.sqlite";
            string connectionString = $"Data Source={databasePath};Version=3;";

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
                        Console.WriteLine($"Number: {number}, Id: {block.Id}, Weight: {block.Weight}, Diameter: {block.Diameter}, Fullname: {block.Fullname}");
                    }
                    else
                    {
                        Console.WriteLine($"Number: {number}, Block with Id '{id}' not found in the database.");
                    }
                }
            }
        }

        public static Block GetBlockInfo(SQLiteConnection connection, string id)
        {
            string selectQuery = "SELECT * FROM Parts WHERE ID = @Id;";

            using (SQLiteCommand command = new SQLiteCommand(selectQuery, connection))
            {
                command.Parameters.AddWithValue("@Id", id);

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Block block = new Block
                        {
                            Id = reader.GetString(reader.GetOrdinal("ID")),
                            Weight = reader.GetDouble(reader.GetOrdinal("Weight")),
                            Diameter = reader.GetInt32(reader.GetOrdinal("Diameter")),
                            Fullname = reader.GetString(reader.GetOrdinal("FullNameTemplate"))
                        };

                        return block;
                    }
                }
            }

            return null;
        }

        public void CreateTableWithFields()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Задаем параметры таблицы
            int numRows = selectedBlockOrderDict.Count; // Количество строк
            int numCols = 4; // Количество столбцов
            double startX = -20000; // Начальная координата X
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

                // Задаем заголовки столбцов
                table.Cells[0, 0].TextString = "№";
                table.Cells[0, 1].TextString = "Наименование";
                table.Cells[0, 2].TextString = "Количество";
                table.Cells[0, 3].TextString = "Масса";

                // Заполняем таблицу данными
                for (int row = 1; row < numRows; row++)
                {
                    for (int col = 0; col < numCols; col++)
                    {
                        Cell cell = table.Cells[row, col];
                        cell.TextHeight = 20;

                        //// Задаем ширину и высоту ячейки
                        //cell.Width = cellWidth;
                        //cell.Height = cellHeight;

                        // Задаем текст для ячейки (пример данных)
                        switch (col)
                        {
                            case 0:
                                cell.TextString = row.ToString();
                                break;
                            case 1:
                                cell.TextString = $"Item {row}";
                                break;
                            case 2:
                                cell.TextString = "10";
                                break;
                            case 3:
                                cell.TextString = "5 kg";
                                break;
                        }
                    }
                }

                tr.Commit();
            }

            ed.Regen(); // Обновляем отображение
        }
    }
}
