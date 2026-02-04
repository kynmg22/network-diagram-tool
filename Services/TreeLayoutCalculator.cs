using System;
using System.Collections.Generic;
using System.Linq;

namespace NetworkDiagramApp
{
    public class TreeLayoutCalculator
    {
        private const int NODE_WIDTH = 120;
        private const int NODE_HEIGHT = 70;
        private const int GAP_X = 30;
        private const int GAP_Y = 80;
        private const int MARGIN_X = 60;
        private const int MARGIN_Y = 40;

        private readonly TreeBuilder _tree;
        private readonly Dictionary<string, double> _subtreeWidths = new();

        public TreeLayoutCalculator(TreeBuilder tree)
        {
            _tree = tree;
        }

        public Dictionary<string, (double X, double Y)> Calculate()
        {
            var positions = new Dictionary<string, (double, double)>();

            if (_tree.Roots.Count == 0)
            {
                throw new Exception("ルートノード（親なしノード）が見つかりません。");
            }

            // 部分木の幅を計算
            foreach (var root in _tree.Roots)
            {
                ComputeSubtreeWidth(root);
            }

            // 位置を割り当て
            if (_tree.Roots.Count == 1)
            {
                AssignPositions(_tree.Roots[0], 0, 0, positions);
            }
            else
            {
                // 複数ルートは横並び
                double currentX = 0;
                foreach (var root in _tree.Roots)
                {
                    AssignPositions(root, currentX, 0, positions);
                    currentX += _subtreeWidths[root] + GAP_X;
                }
            }

            // マージンを適用
            ApplyMargins(positions);

            return positions;
        }

        private double ComputeSubtreeWidth(string nodeID)
        {
            if (_subtreeWidths.ContainsKey(nodeID))
            {
                return _subtreeWidths[nodeID];
            }

            var children = _tree.Children[nodeID];
            if (children.Count == 0)
            {
                _subtreeWidths[nodeID] = NODE_WIDTH;
                return NODE_WIDTH;
            }

            double totalWidth = 0;
            foreach (var child in children)
            {
                totalWidth += ComputeSubtreeWidth(child);
            }
            totalWidth += GAP_X * (children.Count - 1);

            double width = Math.Max(NODE_WIDTH, totalWidth);
            _subtreeWidths[nodeID] = width;
            return width;
        }

        private void AssignPositions(string nodeID, double xLeft, double y, Dictionary<string, (double, double)> positions)
        {
            double myWidth = _subtreeWidths[nodeID];
            double myX = xLeft + (myWidth - NODE_WIDTH) / 2;
            positions[nodeID] = (myX, y);

            var children = _tree.Children[nodeID];
            if (children.Count == 0) return;

            double currentX = xLeft;
            double childY = y + NODE_HEIGHT + GAP_Y;

            foreach (var child in children)
            {
                double childWidth = _subtreeWidths[child];
                AssignPositions(child, currentX, childY, positions);
                currentX += childWidth + GAP_X;
            }
        }

        private void ApplyMargins(Dictionary<string, (double, double)> positions)
        {
            if (positions.Count == 0) return;

            double minX = positions.Values.Min(p => p.Item1);
            double minY = positions.Values.Min(p => p.Item2);

            double shiftX = MARGIN_X - minX;
            double shiftY = MARGIN_Y - minY;

            var keys = positions.Keys.ToList();
            foreach (var key in keys)
            {
                var (x, y) = positions[key];
                positions[key] = (x + shiftX, y + shiftY);
            }
        }
    }
}
