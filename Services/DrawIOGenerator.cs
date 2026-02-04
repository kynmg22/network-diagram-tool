using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace NetworkDiagramApp
{
    public class DrawIOGenerator
    {
        private const int NODE_WIDTH = 120;
        private const int NODE_HEIGHT = 70;
        private const int NOTE_WIDTH = 280;
        private const int NOTE_MIN_HEIGHT = 60;
        private const int NOTE_LINE_HEIGHT = 18;

        private static readonly string[] VLAN_COLORS = new[]
        {
            "#dae8fc", "#d5e8d4", "#fff2cc", "#f8cecc",
            "#e1d5e7", "#d0e0e3", "#fce5cd"
        };

        private readonly Dictionary<string, NetworkNode> _nodes;
        private readonly Dictionary<string, (double X, double Y)> _positions;
        private readonly Dictionary<string, string> _idMap;

        public DrawIOGenerator(Dictionary<string, NetworkNode> nodes, Dictionary<string, (double X, double Y)> positions)
        {
            _nodes = nodes;
            _positions = positions;
            _idMap = new Dictionary<string, string>();

            // IDマッピングを生成
            var usedIDs = new HashSet<string>();
            foreach (var nodeID in nodes.Keys)
            {
                _idMap[nodeID] = ToSafeID(nodeID, usedIDs);
            }
        }

        public void Generate(string outputPath)
        {
            var vlanGroups = GroupByVLAN();

            var framesXML = GenerateVLANFrames(vlanGroups);
            var cellsXML = GenerateNodeCells();
            var edgesXML = GenerateEdges();
            var noteXML = GenerateNoteBox();

            var xmlContent = BuildXMLDocument(framesXML, cellsXML, edgesXML, noteXML);

            // UTF-8 BOMなしで保存
            var utf8WithoutBom = new UTF8Encoding(false);
            File.WriteAllText(outputPath, xmlContent, utf8WithoutBom);
        }

        private Dictionary<int, List<string>> GroupByVLAN()
        {
            var groups = new Dictionary<int, List<string>>();
            foreach (var kvp in _nodes)
            {
                if (kvp.Value.VLAN.HasValue)
                {
                    int vlan = kvp.Value.VLAN.Value;
                    if (!groups.ContainsKey(vlan))
                    {
                        groups[vlan] = new List<string>();
                    }
                    groups[vlan].Add(kvp.Key);
                }
            }
            return groups;
        }

        private List<string> GenerateVLANFrames(Dictionary<int, List<string>> vlanGroups)
        {
            var frames = new List<string>();
            var manager = new VLANFrameManager();

            foreach (var vlan in vlanGroups.Keys.OrderBy(v => v))
            {
                var bounds = manager.CalculateFrameBounds(vlanGroups[vlan], _positions);
                if (bounds == null) continue;

                string color = VLAN_COLORS[(vlan - 1) % VLAN_COLORS.Length];
                string style = $"rounded=0;html=1;whiteSpace=wrap;fillColor={color};fillOpacity=15;" +
                               "strokeColor=#666666;strokeOpacity=80;align=right;verticalAlign=top;" +
                               "spacingRight=6;spacingTop=6;";

                frames.Add(
                    $"        <mxCell id=\"frame_{vlan}\" value=\"VLAN{vlan}\" style=\"{style}\" vertex=\"1\" parent=\"1\">\n" +
                    $"          <mxGeometry x=\"{(int)bounds.MinX}\" y=\"{(int)bounds.MinY}\" " +
                    $"width=\"{(int)(bounds.MaxX - bounds.MinX)}\" height=\"{(int)(bounds.MaxY - bounds.MinY)}\" as=\"geometry\"/>\n" +
                    $"        </mxCell>");
            }

            return frames;
        }

        private List<string> GenerateNodeCells()
        {
            var cells = new List<string>();

            foreach (var kvp in _nodes)
            {
                string nodeID = kvp.Key;
                var node = kvp.Value;
                string drawID = _idMap[nodeID];

                var (x, y) = _positions.ContainsKey(nodeID) ? _positions[nodeID] : (60, 40);

                string label = MakeNodeLabel(node);
                string style = StyleNode(node);

                cells.Add(
                    $"        <mxCell id=\"{drawID}\" value=\"{label}\" style=\"{style}\" vertex=\"1\" parent=\"1\">\n" +
                    $"          <mxGeometry x=\"{(int)x}\" y=\"{(int)y}\" width=\"{NODE_WIDTH}\" height=\"{NODE_HEIGHT}\" as=\"geometry\"/>\n" +
                    $"        </mxCell>");
            }

            return cells;
        }

        private List<string> GenerateEdges()
        {
            var edges = new List<string>();
            int edgeCounter = 1;

            foreach (var kvp in _nodes)
            {
                string childID = kvp.Key;
                var node = kvp.Value;

                int numParents = node.Parents.Count;
                if (numParents == 0) continue;

                for (int i = 0; i < numParents; i++)
                {
                    string parentID = node.Parents[i];
                    double exitX = 0.5;
                    double entryX = numParents == 1 ? 0.5 : (i + 1.0) / (numParents + 1.0);

                    string style = $"edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;" +
                                   $"jettySize=auto;html=1;endArrow=none;" +
                                   $"exitX={exitX};exitY=1;exitDx=0;exitDy=0;" +
                                   $"entryX={entryX};entryY=0;entryDx=0;entryDy=0;";

                    edges.Add(
                        $"        <mxCell id=\"e{edgeCounter}\" style=\"{style}\" edge=\"1\" parent=\"1\" " +
                        $"source=\"{_idMap[parentID]}\" target=\"{_idMap[childID]}\">\n" +
                        $"          <mxGeometry relative=\"1\" as=\"geometry\"/>\n" +
                        $"        </mxCell>");

                    edgeCounter++;
                }
            }

            return edges;
        }

        private string GenerateNoteBox()
        {
            var noteItems = _nodes.Values
                .Where(n => !string.IsNullOrWhiteSpace(n.Note))
                .OrderBy(n => n.ExcelRow)
                .ToList();

            if (!noteItems.Any()) return string.Empty;

            var lines = new List<string>();
            for (int i = 0; i < noteItems.Count; i++)
            {
                if (i > 0)
                {
                    lines.Add("");
                }
                lines.Add($"■ {noteItems[i].Name}");
                lines.Add(noteItems[i].Note);
            }

            string plainText = string.Join("\n", lines);
            string htmlText = TextToDrawIOHTML(plainText);

            int lineCount = htmlText.Split(new[] { "&lt;br&gt;" }, StringSplitOptions.None).Length;
            int noteHeight = Math.Max(NOTE_MIN_HEIGHT, lineCount * NOTE_LINE_HEIGHT + 20);

            double maxX = _positions.Values.Max(p => p.X) + NODE_WIDTH;
            double minY = _positions.Values.Min(p => p.Y);
            double noteX = maxX + 50;
            double noteY = minY;

            string style = "rounded=0;html=1;whiteSpace=wrap;align=left;verticalAlign=top;" +
                           "fillColor=#ffffcc;strokeColor=#666666;" +
                           "spacingTop=10;spacingLeft=10;spacingRight=10;spacingBottom=10;fontSize=11;";

            return $"        <mxCell id=\"note_box_1\" value=\"{htmlText}\" style=\"{style}\" vertex=\"1\" parent=\"1\">\n" +
                   $"          <mxGeometry x=\"{(int)noteX}\" y=\"{(int)noteY}\" width=\"{NOTE_WIDTH}\" height=\"{noteHeight}\" as=\"geometry\"/>\n" +
                   $"        </mxCell>";
        }

        private string BuildXMLDocument(List<string> frames, List<string> cells, List<string> edges, string note)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<mxfile host=\"app.diagrams.net\" modified=\"2024-01-01T00:00:00.000Z\" agent=\"Network Diagram Generator\" version=\"22.0.0\" etag=\"generated\" type=\"device\">");
            sb.AppendLine("  <diagram name=\"Network\" id=\"network-diagram\">");
            sb.AppendLine("    <mxGraphModel dx=\"1000\" dy=\"1000\" grid=\"1\" gridSize=\"10\" guides=\"1\" tooltips=\"1\" connect=\"1\" arrows=\"1\" fold=\"1\" page=\"1\" pageScale=\"1\" pageWidth=\"827\" pageHeight=\"1169\" math=\"0\" shadow=\"0\">");
            sb.AppendLine("      <root>");
            sb.AppendLine("        <mxCell id=\"0\"/>");
            sb.AppendLine("        <mxCell id=\"1\" parent=\"0\"/>");

            foreach (var frame in frames)
            {
                sb.AppendLine(frame);
            }

            foreach (var cell in cells)
            {
                sb.AppendLine(cell);
            }

            foreach (var edge in edges)
            {
                sb.AppendLine(edge);
            }

            if (!string.IsNullOrEmpty(note))
            {
                sb.AppendLine(note);
            }

            sb.AppendLine("      </root>");
            sb.AppendLine("    </mxGraphModel>");
            sb.AppendLine("  </diagram>");
            sb.AppendLine("</mxfile>");

            return sb.ToString();
        }

        private static string ToSafeID(string excelID, HashSet<string> usedIDs)
        {
            string s = excelID.Trim();
            s = s.Replace(" ", "_").Replace("\t", "_");
            s = s.Replace("<", "_").Replace(">", "_");
            s = s.Replace("\"", "_").Replace("'", "_");
            s = s.Replace("&", "_");
            s = Regex.Replace(s, "_+", "_").Trim('_');

            if (string.IsNullOrEmpty(s)) s = "node";

            string baseID = "n_" + s;
            string resultID = baseID;
            int counter = 1;

            while (usedIDs.Contains(resultID))
            {
                resultID = $"{baseID}_{counter}";
                counter++;
            }

            usedIDs.Add(resultID);
            return resultID;
        }

        private static string MakeNodeLabel(NetworkNode node)
        {
            string id = HttpUtility.HtmlEncode(node.ID);
            string name = HttpUtility.HtmlEncode(node.Name);
            string ip = HttpUtility.HtmlEncode(node.IP);

            if (!string.IsNullOrEmpty(ip))
            {
                return $"{id}&lt;br&gt;{name}&lt;br&gt;{ip}";
            }
            return $"{id}&lt;br&gt;{name}";
        }

        private static string TextToDrawIOHTML(string text)
        {
            string escaped = HttpUtility.HtmlEncode(text);
            return escaped.Replace("\n", "&lt;br&gt;");
        }

        private static string StyleNode(NetworkNode node)
        {
            string baseStyle = "rounded=1;html=1;whiteSpace=wrap;align=center;verticalAlign=middle;";

            if (node.IsONU())
            {
                return baseStyle + "fillColor=#cfe2f3;strokeColor=#1c4587;";
            }
            return baseStyle + "fillColor=#f5f5f5;strokeColor=#666666;";
        }
    }
}
