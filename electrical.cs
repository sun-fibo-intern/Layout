using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Connection
{
    public class SLDGenerator
    {
       
        // =========================================================
        // MODELS
        // =========================================================

        public class Node
        {
            public int Id { get; set; }

            public Point3d Position { get; set; }

            public List<int> ConnectedNodes { get; set; }
                = new List<int>();
        }

        public class Wire
        {
            public int StartNode { get; set; }

            public int EndNode { get; set; }

            public double Length { get; set; }
        }

        public class LoadData
        {
            public string CircuitNo { get; set; }

            public string DeviceType { get; set; }

            public string NodeName { get; set; }

            public double Watt { get; set; }

            public string Phase { get; set; }
            public string Code { get; set; }
        }

        public class Circuit
        {
            public string CircuitNo { get; set; }

            public string Description { get; set; }

            public string Phase { get; set; }

            public double TotalLoad { get; set; }

            public List<string> ConnectedNodes
                = new List<string>();
            public List<string> ExcelNodes
    = new List<string>();
        }

        // =========================================================
        // MAIN COMMAND
        // =========================================================

        [CommandMethod("GENERATE_SLD")]
        public void GenerateSLD()
        {
            Document doc =
                Application.DocumentManager.MdiActiveDocument;

            Editor ed = doc.Editor;

            Database db = doc.Database;

            try
            {
                // =================================================
                // STEP 1 : SELECT EXCEL FILE
                // =================================================

                string excelPath =
                    SelectExcelFile(ed);

                if (string.IsNullOrEmpty(excelPath))
                    return;

                // =================================================
                // STEP 2 : GET SHEET NAMES
                // =================================================

                List<string> sheetNames =
                    GetSheetNames(excelPath);

                if (sheetNames.Count == 0)
                {
                    ed.WriteMessage(
                        "\nNo sheets found.");

                    return;
                }

                // =================================================
                // STEP 3 : ASK USER TO SELECT SHEET
                // =================================================

                string selectedSheet =
                    AskUserToSelectSheet(
                        ed,
                        sheetNames);

                if (string.IsNullOrEmpty(selectedSheet))
                    return;

                // =================================================
                // STEP 4 : READ EXCEL DATA
                // =================================================

                List<LoadData> excelLoads =
                    ReadLoadData(
                        excelPath,
                        selectedSheet);

                List<LoadData> wireNodeMappings =
    ReadWireNodeMapping(
        excelPath,
        selectedSheet);

                var circuitList =
    wireNodeMappings
    .Select(x => x.CircuitNo)
    .Distinct()
    .ToList();

                ed.WriteMessage(
                    $"\nUnique Circuits Found : {circuitList.Count}");

                foreach (string c in circuitList)
                {
                    ed.WriteMessage($"\nCircuit : {c}");
                }

                ed.WriteMessage(
                    $"\nWire-Node Mappings : {wireNodeMappings.Count}");

                foreach (var item in wireNodeMappings.Take(20))
                {
                    ed.WriteMessage(
                        $"\nCircuit={item.CircuitNo}  Node={item.NodeName}");
                }

                ed.WriteMessage(
                    $"\nLoaded {excelLoads.Count} rows.");

                // =================================================
                // STEP 5 : SELECT DWG AREA
                // =================================================

                SelectionSet selection =
                    SelectDWGWindow(ed);

                if (selection == null)
                    return;

                using (Transaction tr =
                    db.TransactionManager.StartTransaction())
                {
                    // =============================================
                    // STEP 6 : EXTRACT NODES
                    // =============================================

                    List<Node> nodes =
                        ExtractNodes(
                            selection,
                            tr);

                    ed.WriteMessage(
                        $"\nNodes Found : {nodes.Count}");

                    // =============================================
                    // STEP 7 : BUILD WIRES
                    // =============================================

                    List<Wire> wires =
                        BuildWireConnections(
                            selection,
                            nodes,
                            tr);

                    ed.WriteMessage(
                        $"\nWires Found : {wires.Count}");

                    // =============================================
                    // STEP 8 : BUILD GRAPH
                    // =============================================

                    Dictionary<int, List<int>> graph =
                        BuildGraph(wires);

                    // =============================================
                    // STEP 9 : GREEDY ROUTING
                    // =============================================

                    List<List<int>> paths =
                        GenerateGreedyPaths(
                            graph,
                            nodes);

                    ed.WriteMessage(
                        $"\nPaths Generated : {paths.Count}");

                    // =============================================
                    // STEP 10 : BUILD CIRCUITS
                    // =============================================

                    List<Circuit> circuits = BuildCircuits(paths, excelLoads, wireNodeMappings);

                    // =============================================
                    // STEP 11 : ASK USER FOR INSERTION POINT
                    // =============================================

                    Point3d? insertionPoint =
                        GetInsertionPoint(ed);

                    if (insertionPoint == null)
                    {
                        ed.WriteMessage(
                            "\nNo insertion point selected.");

                        return;
                    }



                    // =============================================
                    // STEP 12 : DRAW SLD
                    // =============================================

                    DrawSLD(
                        db,
                        tr,
                        circuits,
                        insertionPoint.Value);

                    tr.Commit();
                }


                ed.WriteMessage(
                    "\nSLD Generated Successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(
                    $"\nERROR : {ex.Message}");
            }
        }

        public void DrawBranch(
    BlockTableRecord ms,
    Transaction tr,
    Point3d start,
    string nodeName)
        {
            double branchLength = 80;

            Point3d end =
                new Point3d(
                    start.X - branchLength,
                    start.Y,
                    0);

            Line branch =
                new Line(start, end);

            ms.AppendEntity(branch);
            tr.AddNewlyCreatedDBObject(branch, true);

            DBText txt =
                new DBText();

            txt.Position =
                new Point3d(
                    end.X - 20,
                    end.Y - 5,
                    0);

            txt.Height = 8;

            txt.TextString = nodeName;

            ms.AppendEntity(txt);
            tr.AddNewlyCreatedDBObject(txt, true);
        }

        // =========================================================
        // METHOD 1 : SELECT EXCEL FILE
        // =========================================================

        public string SelectExcelFile(Editor ed)
        {
            PromptOpenFileOptions options =
                new PromptOpenFileOptions(
                    "\nSelect Excel File");

            options.Filter =
                "Excel Files (*.xlsx)|*.xlsx";

            PromptFileNameResult result =
                ed.GetFileNameForOpen(options);

            if (result.Status != PromptStatus.OK)
                return null;

            return result.StringResult;
        }

        // =========================================================
        // METHOD 2 : GET SHEET NAMES
        // =========================================================

        public List<string> GetSheetNames(
     string excelPath)
        {
            List<string> sheets =
                new List<string>();

            string connString =
                @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=" + excelPath + ";" +
                "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

            using (OleDbConnection conn =
                new OleDbConnection(connString))
            {
                conn.Open();

                System.Data.DataTable dt =
    conn.GetOleDbSchemaTable(
        OleDbSchemaGuid.Tables,
        null);

                foreach (DataRow row in dt.Rows)
                {
                    string sheetName =
                        row["TABLE_NAME"]
                        .ToString();

                    Application.DocumentManager
                        .MdiActiveDocument.Editor
                        .WriteMessage(
                            $"\nRAW SHEET = [{sheetName}]");

                    // Keep only real worksheet names
                    if (!sheetName.EndsWith("$'") &&
                        !sheetName.EndsWith("$"))
                    {
                        continue;
                    }

                    sheetName =
                        sheetName
                        .Replace("'", "")
                        .Replace("$", "")
                        .Trim();

                    if (!sheets.Contains(sheetName))
                    {
                        sheets.Add(sheetName);
                    }
                }
            }

            return sheets;
        }

        // =========================================================
        // METHOD 3 : ASK USER TO SELECT SHEET
        // =========================================================

        public string AskUserToSelectSheet(
     Editor ed,
     List<string> sheetNames)
        {
            ed.WriteMessage(
                "\nAvailable Sheets:\n");

            for (int i = 0;
                i < sheetNames.Count;
                i++)
            {
                ed.WriteMessage(
                    $"\n{i + 1} : {sheetNames[i]}");
            }

            PromptIntegerOptions options =
                new PromptIntegerOptions(
                    "\nEnter Sheet Number : ");

            options.AllowNegative = false;
            options.AllowZero = false;

            options.LowerLimit = 1;

            options.UpperLimit =
                sheetNames.Count;

            PromptIntegerResult result =
                ed.GetInteger(options);

            if (result.Status != PromptStatus.OK)
                return null;

            int selectedIndex =
                result.Value - 1;

            return sheetNames[selectedIndex];
        }

        // =========================================================
        // METHOD 4 : READ EXCEL DATA
        // =========================================================

        public List<LoadData> ReadLoadData(
    string excelPath,
    string sheetName)
        {
            List<LoadData> loads =
                new List<LoadData>();

            string connString =
                @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=" + excelPath + ";" +
                "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

            using (OleDbConnection conn =
                new OleDbConnection(connString))
            {
                conn.Open();

                string query =
                    $"SELECT * FROM [{sheetName}$]";

                OleDbDataAdapter adapter =
                    new OleDbDataAdapter(
                        query,
                        conn);

                System.Data.DataTable dt =
    new System.Data.DataTable();

                adapter.Fill(dt);

                Editor ed =
    Application.DocumentManager
    .MdiActiveDocument.Editor;

                ed.WriteMessage(
                    $"\n====================");

                ed.WriteMessage(
                    $"\nROWS FOUND : {dt.Rows.Count}");

                ed.WriteMessage(
                    $"\nCOLUMNS FOUND : {dt.Columns.Count}");

                for (int i = 0;
                     i < 15 && i < dt.Rows.Count;
                     i++)
                {
                    string firstCol =
                        dt.Rows[i][0]?.ToString();

                    string secondCol =
                        dt.Rows[i][1]?.ToString();

                    string thirdCol =
                        dt.Rows[i][2]?.ToString();

                    string fourthCol =
                        dt.Rows[i][3]?.ToString();

                    ed.WriteMessage(
                        $"\nROW {i} => " +
                        $"{firstCol} | " +
                        $"{secondCol} | " +
                        $"{thirdCol} | " +
                        $"{fourthCol}");
                }

                ed.WriteMessage(
                    $"\n====================");

                foreach (DataRow row in dt.Rows)
                {
                    LoadData data =
                        new LoadData();

                    data.Code =
                        row[0]?.ToString();

                    double watt = 0;

                    double.TryParse(
                        row[1]?.ToString(),
                        out watt);

                    data.Watt = watt;

                    data.Phase =
                        row[2]?.ToString();

                    loads.Add(data);
                }
            }

            return loads;
        }
        private List<string> SplitNodeNames(string value)
        {
            List<string> nodes =
                new List<string>();

            if (string.IsNullOrWhiteSpace(value))
                return nodes;

            string[] parts =
                value.Split(',');

            foreach (string part in parts)
            {
                string node =
                    part.Trim();

                if (!string.IsNullOrWhiteSpace(node))
                {
                    nodes.Add(node);
                }
            }

            return nodes;
        }
        public List<LoadData> ReadWireNodeMapping(
     string excelPath,
     string sheetName)
        {
            List<LoadData> mappings =
                new List<LoadData>();
            Editor ed =
    Application.DocumentManager
    .MdiActiveDocument.Editor;

            string connString =
                @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=" + excelPath + ";" +
                "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

            using (OleDbConnection conn =
                new OleDbConnection(connString))
            {
                conn.Open();

                string query =
                    $"SELECT * FROM [{sheetName}$]";

                OleDbDataAdapter adapter =
                    new OleDbDataAdapter(
                        query,
                        conn);

                System.Data.DataTable dt =
    new System.Data.DataTable();

                adapter.Fill(dt);

                if (dt.Rows.Count < 10)
                    return mappings;

                int totalCols =
                    dt.Columns.Count;

                ed.WriteMessage(
    "\nHEADER ROW CHECK");

                for (int r = 0; r < 12; r++)
                {
                    ed.WriteMessage($"\n\nROW {r}");

                    for (int c = 0; c < totalCols; c++)
                    {
                        string value =
                            dt.Rows[r][c]?.ToString();

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            ed.WriteMessage(
                                $"\nCOL {c} = {value}");
                        }
                    }
                }

                for (int row = 8;
                     row < dt.Rows.Count;
                     row++)
                {
                    string circuitNo =
                        dt.Rows[row][3]
                        ?.ToString();

                    if (string.IsNullOrWhiteSpace(
                        circuitNo))
                    {
                        continue;
                    }

                    for (int col = 8;
                         col < totalCols;
                         col++)
                    {
                        string nodeValue =
                            dt.Rows[row][col]
                            ?.ToString();

                        if (string.IsNullOrWhiteSpace(
                            nodeValue))
                        {
                            continue;
                        }

                        List<string> nodes =
                            SplitNodeNames(
                                nodeValue);

                        foreach (string node in nodes)
                        {
                            LoadData item =
                                new LoadData
                                {
                                    CircuitNo = circuitNo,
                                    NodeName = node,
                                    DeviceType =
                                        dt.Rows[6][col]
                                        ?.ToString()
                                };

                            mappings.Add(item);

                            ed.WriteMessage(
                                $"\nMAPPING -> Circuit={item.CircuitNo}  Node={item.NodeName}");
                        }
                    }
                }
            }

            return mappings;
        }
        // =========================================================
        // METHOD 5 : SELECT DWG WINDOW
        // =========================================================

        public SelectionSet SelectDWGWindow(Editor ed)
        {
            PromptPointResult firstCorner =
                ed.GetPoint(
                    "\nPick first corner of DWG area : ");

            if (firstCorner.Status != PromptStatus.OK)
                return null;

            PromptCornerOptions cornerOptions =
                new PromptCornerOptions(
                    "\nPick opposite corner : ",
                    firstCorner.Value);

            PromptPointResult secondCorner =
                ed.GetCorner(cornerOptions);

            if (secondCorner.Status != PromptStatus.OK)
                return null;

            PromptSelectionResult result =
                ed.SelectWindow(
                    firstCorner.Value,
                    secondCorner.Value);

            if (result.Status != PromptStatus.OK)
                return null;

            int entityCount =
                result.Value.Count;

            ed.WriteMessage(
                $"\nTotal Entities Selected : {entityCount}");

            return result.Value;
        }

        public Point3d? GetInsertionPoint(Editor ed)
        {
            PromptPointOptions ppo =
                new PromptPointOptions(
                    "\nPick insertion point for SLD : ");

            PromptPointResult ppr =
                ed.GetPoint(ppo);

            if (ppr.Status != PromptStatus.OK)
                return null;

            return ppr.Value;
        }
        public List<Node> ExtractNodes(
    SelectionSet selection,
    Transaction tr)
        {
            List<Node> nodes =
                new List<Node>();

            int nodeId = 1;

            foreach (SelectedObject obj in selection)
            {
                if (obj == null)
                    continue;

                Entity ent =
                    tr.GetObject(
                        obj.ObjectId,
                        OpenMode.ForRead)
                    as Entity;

                // LINE SUPPORT

                if (ent is Line line)
                {
                    nodes.Add(new Node
                    {
                        Id = nodeId++,
                        Position = line.StartPoint
                    });

                    nodes.Add(new Node
                    {
                        Id = nodeId++,
                        Position = line.EndPoint
                    });
                }

                // POLYLINE SUPPORT

                else if (ent is Polyline pline)
                {
                    for (int i = 0;
                        i < pline.NumberOfVertices;
                        i++)
                    {
                        nodes.Add(new Node
                        {
                            Id = nodeId++,
                            Position =
                                pline.GetPoint3dAt(i)
                        });
                    }
                }
            }

            return RemoveDuplicateNodes(nodes);
        }
        public List<Node> RemoveDuplicateNodes(
    List<Node> nodes)
        {
            List<Node> cleanNodes =
                new List<Node>();

            double tolerance = 5;

            foreach (Node node in nodes)
            {
                bool exists =
                    cleanNodes.Any(x =>
                        x.Position.DistanceTo(
                            node.Position)
                        < tolerance);

                if (!exists)
                {
                    cleanNodes.Add(node);
                }
            }

            return cleanNodes;
        }
        public List<Wire> BuildWireConnections(
    SelectionSet selection,
    List<Node> nodes,
    Transaction tr)
        {
            List<Wire> wires =
                new List<Wire>();

            foreach (SelectedObject obj in selection)
            {
                if (obj == null)
                    continue;

                Entity ent =
                    tr.GetObject(
                        obj.ObjectId,
                        OpenMode.ForRead)
                    as Entity;

                // LINE SUPPORT

                if (ent is Line line)
                {
                    Node startNode =
                        GetNearestNode(
                            nodes,
                            line.StartPoint);

                    Node endNode =
                        GetNearestNode(
                            nodes,
                            line.EndPoint);

                    if (startNode != null &&
                        endNode != null)
                    {
                        wires.Add(new Wire
                        {
                            StartNode = startNode.Id,

                            EndNode = endNode.Id,

                            Length =
                                line.StartPoint.DistanceTo(
                                    line.EndPoint)
                        });
                    }
                }

                // POLYLINE SUPPORT

                else if (ent is Polyline pline)
                {
                    for (int i = 0;
                        i < pline.NumberOfVertices - 1;
                        i++)
                    {
                        Point3d p1 =
                            pline.GetPoint3dAt(i);

                        Point3d p2 =
                            pline.GetPoint3dAt(i + 1);

                        Node startNode =
                            GetNearestNode(nodes, p1);

                        Node endNode =
                            GetNearestNode(nodes, p2);

                        if (startNode != null &&
                            endNode != null)
                        {
                            wires.Add(new Wire
                            {
                                StartNode = startNode.Id,

                                EndNode = endNode.Id,

                                Length =
                                    p1.DistanceTo(p2)
                            });
                        }
                    }
                }
            }

            return wires;
        }
        public Node GetNearestNode(
    List<Node> nodes,
    Point3d point)
        {
            double minDistance =
                double.MaxValue;

            Node nearest = null;

            foreach (Node node in nodes)
            {
                double distance =
                    node.Position.DistanceTo(point);

                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearest = node;
                }
            }

            return nearest;
        }
        public Dictionary<int, List<int>> BuildGraph(
    List<Wire> wires)
        {
            Dictionary<int, List<int>> graph =
                new Dictionary<int, List<int>>();

            foreach (Wire wire in wires)
            {
                if (!graph.ContainsKey(
                    wire.StartNode))
                {
                    graph[wire.StartNode] =
                        new List<int>();
                }

                if (!graph.ContainsKey(
                    wire.EndNode))
                {
                    graph[wire.EndNode] =
                        new List<int>();
                }

                graph[wire.StartNode]
                    .Add(wire.EndNode);

                graph[wire.EndNode]
                    .Add(wire.StartNode);
            }

            return graph;
        }
        public List<List<int>> GenerateGreedyPaths(
    Dictionary<int, List<int>> graph,
    List<Node> nodes)
        {
            List<List<int>> paths =
                new List<List<int>>();

            HashSet<int> visited =
                new HashSet<int>();

            foreach (Node startNode in nodes)
            {
                if (visited.Contains(startNode.Id))
                    continue;

                List<int> currentPath =
                    new List<int>();

                int current =
                    startNode.Id;

                while (true)
                {
                    currentPath.Add(current);

                    visited.Add(current);

                    if (!graph.ContainsKey(current))
                        break;

                    List<int> neighbors =
                        graph[current]
                        .Where(x =>
                            !visited.Contains(x))
                        .ToList();

                    if (neighbors.Count == 0)
                        break;

                    int next =
                        neighbors.First();

                    current = next;
                }

                paths.Add(currentPath);
            }

            return paths;
        }
        public List<Circuit> BuildCircuits(
     List<List<int>> paths,
     List<LoadData> excelLoads,
     List<LoadData> wireNodeMappings)
        {
            List<Circuit> circuits =
                new List<Circuit>();

            int circuitCounter = 1;

            List<string> excelCircuits =
                wireNodeMappings
                .Select(x => x.CircuitNo)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .ToList();

            foreach (List<int> path in paths)
            {
                Circuit circuit =
                    new Circuit();

                string circuitRef;

                if (circuitCounter <= excelCircuits.Count)
                {
                    circuitRef =
                        excelCircuits[circuitCounter - 1];
                }
                else
                {
                    circuitRef =
                        $"C{circuitCounter}";
                }

                circuit.CircuitNo =
                    circuitRef;

                var mappedNodes =
     wireNodeMappings
     .Where(x =>
         x.CircuitNo.Trim().Equals(
             circuitRef.Trim(),
             StringComparison.OrdinalIgnoreCase))
     .Select(x => x.NodeName)
     .Distinct()
     .ToList();

                circuit.ExcelNodes.AddRange(mappedNodes);

                foreach (string node in mappedNodes)
                {
                    Application.DocumentManager
                        .MdiActiveDocument.Editor
                        .WriteMessage(
                            $"\n{circuitRef} -> {node}");
                }

                Application.DocumentManager
    .MdiActiveDocument.Editor
    .WriteMessage(
        $"\nCircuit {circuitRef}");

                foreach (string node in mappedNodes)
                {
                    Application.DocumentManager
                        .MdiActiveDocument.Editor
                        .WriteMessage(
                            $"\n   Node -> {node}");
                }

                Application.DocumentManager
                    .MdiActiveDocument.Editor
                    .WriteMessage(
                        $"\n{circuitRef} -> {mappedNodes.Count} mapped nodes");

                if (circuitCounter <= excelLoads.Count)
                {
                    LoadData load =
                        excelLoads[circuitCounter - 1];

                    circuit.Description =
                        load.Code;

                    circuit.TotalLoad =
                        load.Watt;

                    circuit.Phase =
                        load.Phase;
                }
                else
                {
                    circuit.Description =
                        "UNASSIGNED";

                    circuit.TotalLoad = 0;

                    circuit.Phase = "R";
                }

                foreach (int nodeId in path)
                {
                    circuit.ConnectedNodes
                        .Add(nodeId.ToString());
                }

                circuits.Add(circuit);

                circuitCounter++;
            }

            return circuits;
        }
        public void DrawSLD(
    Database db,
    Transaction tr,
    List<Circuit> circuits,
    Point3d insertionPoint)
        {
            BlockTable bt =
                tr.GetObject(
                    db.BlockTableId,
                    OpenMode.ForRead)
                as BlockTable;

            BlockTableRecord ms =
                tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite)
                as BlockTableRecord;

            // =============================================
            // SPACING SETTINGS
            // =============================================

            double delta = 150;

            double phaseGap = delta * 3;

            double currentX =
                insertionPoint.X;

            double fixedY =
                insertionPoint.Y;

            var rCircuits =
    circuits
    .Where(c => c.CircuitNo.StartsWith("R"))
    .ToList();

            var yCircuits =
                circuits
                .Where(c => c.CircuitNo.StartsWith("Y"))
                .ToList();

            var bCircuits =
                circuits
                .Where(c => c.CircuitNo.StartsWith("B"))
                .ToList();

            List<Circuit> orderedCircuits =
                new List<Circuit>();

            orderedCircuits.AddRange(rCircuits);

            orderedCircuits.AddRange(yCircuits);

            orderedCircuits.AddRange(bCircuits);

            foreach (Circuit circuit in orderedCircuits)
            {

                Application.DocumentManager
    .MdiActiveDocument.Editor
    .WriteMessage(
        $"\nDRAWING CIRCUIT : {circuit.CircuitNo}");

                Application.DocumentManager
                    .MdiActiveDocument.Editor
                    .WriteMessage(
                        $"\nEXCEL NODES COUNT : {circuit.ExcelNodes.Count}");
                // =========================================
                // MAIN HORIZONTAL LINE
                // =========================================

                Point3d startPoint =
   new Point3d(
       currentX,
       fixedY,
       0);
                Point3d endPoint =
  new Point3d(
      currentX,
      fixedY,
      0);
                short colorIndex = 7; // white default

                if (circuit.CircuitNo.StartsWith("R"))
                {
                    colorIndex = 1; // Red
                }
                else if (circuit.CircuitNo.StartsWith("Y"))
                {
                    colorIndex = 2; // Yellow
                }
                else if (circuit.CircuitNo.StartsWith("B"))
                {
                    colorIndex = 5; // Blue
                }

                Line wire =
                    new Line(
                        startPoint,
                        endPoint);

                wire.ColorIndex = colorIndex;

                ms.AppendEntity(wire);

                tr.AddNewlyCreatedDBObject(
                    wire,
                    true);

                // =========================================
                // NODE CIRCLE
                // =========================================

                Circle node =
                    new Circle();

                node.Center = startPoint;

                node.Radius = 5;

                node.ColorIndex = colorIndex;

                ms.AppendEntity(node);

                tr.AddNewlyCreatedDBObject(
                    node,
                    true);

                // =========================================
                // CIRCUIT TEXT
                // =========================================

                DBText text =
                    new DBText();

                text.Position =
      new Point3d(
          currentX + 20,
          fixedY,
          0);

                text.Height = 10;

                text.ColorIndex = colorIndex;

                text.TextString =
                    $"{circuit.CircuitNo} | " +
                    $"{circuit.Description} | " +
                    $"{circuit.Phase} | " +
                    $"{circuit.TotalLoad}W";

                ms.AppendEntity(text);

                tr.AddNewlyCreatedDBObject(
                    text,
                    true);

                // =========================================
                // CONNECTED NODE TEXT
                // =========================================
                DrawConnectedLoads(
                 ms,
                 tr,
                 endPoint,
                 circuit.ExcelNodes,
                 colorIndex);

                // =========================================
                // MOVE NEXT CIRCUIT DOWN
                // =========================================
                currentX += delta;

                if (circuit.CircuitNo.StartsWith("R"))
                {
                    bool isLastR =
                        circuit ==
                        rCircuits.Last();

                    if (isLastR)
                    {
                        currentX += phaseGap;
                    }
                }

                if (circuit.CircuitNo.StartsWith("Y"))
                {
                    bool isLastY =
                        circuit ==
                        yCircuits.Last();

                    if (isLastY)
                    {
                        currentX += phaseGap;
                    }
                }
            }
        }

        public void DrawConnectedLoads(
     BlockTableRecord ms,
     Transaction tr,
     Point3d circuitPoint,
     List<string> loads,
     short colorIndex)
        {
            if (loads.Count == 0)
                return;

            double busHeight =
                loads.Count * 25;

            double busX =
                circuitPoint.X;

            Point3d busTop =
                new Point3d(
                    busX,
                    circuitPoint.Y,
                    0);

            Point3d busBottom =
                new Point3d(
                    busX,
                    circuitPoint.Y - busHeight,
                    0);

            Line verticalBus =
                new Line(
                    busTop,
                    busBottom);

            verticalBus.ColorIndex =
                colorIndex;

            ms.AppendEntity(verticalBus);

            tr.AddNewlyCreatedDBObject(
                verticalBus,
                true);

            for (int i = 0; i < loads.Count; i++)
            {
                double y =
                    circuitPoint.Y -
                    ((i + 1) * 25);

                Point3d left =
                    new Point3d(
                        busX - 60,
                        y,
                        0);

                Point3d right =
                    new Point3d(
                        busX,
                        y,
                        0);

                Line branch =
                    new Line(
                        left,
                        right);

                branch.ColorIndex =
                    colorIndex;

                ms.AppendEntity(branch);

                tr.AddNewlyCreatedDBObject(
                    branch,
                    true);

                DBText txt =
                    new DBText();

                txt.Position =
                    new Point3d(
                        left.X - 40,
                        y - 3,
                        0);

                txt.Height = 8;

                txt.ColorIndex =
                    colorIndex;

                txt.TextString =
                    loads[i];

                ms.AppendEntity(txt);

                tr.AddNewlyCreatedDBObject(
                    txt,
                    true);
            }
        }

    }
}
