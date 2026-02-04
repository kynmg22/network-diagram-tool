using System;
using System.Collections.Generic;
using System.Linq;

namespace NetworkDiagramApp
{
    public class VLANFrameManager
    {
        private const int NODE_WIDTH = 120;
        private const int NODE_HEIGHT = 70;
        private const int FRAME_PAD_X = 25;
        private const int FRAME_PAD_TOP = 45;
        private const int FRAME_PAD_BOTTOM = 25;
        private const int VLAN_FRAME_GAP = 40;

        public class Bounds
        {
            public double MinX { get; set; }
            public double MinY { get; set; }
            public double MaxX { get; set; }
            public double MaxY { get; set; }

            public double CenterX => (MinX + MaxX) / 2;

            public bool OverlapsVertically(Bounds other)
            {
                return !(MaxY < other.MinY || other.MaxY < MinY);
            }
        }

        public Bounds? CalculateFrameBounds(List<string> nodeIDs, Dictionary<string, (double X, double Y)> positions)
        {
            var validPositions = nodeIDs
                .Where(nid => positions.ContainsKey(nid))
                .Select(nid => positions[nid])
                .ToList();

            if (!validPositions.Any()) return null;

            var xs = validPositions.Select(p => p.X).ToList();
            var ys = validPositions.Select(p => p.Y).ToList();

            return new Bounds
            {
                MinX = xs.Min() - FRAME_PAD_X,
                MinY = ys.Min() - FRAME_PAD_TOP,
                MaxX = xs.Max() + NODE_WIDTH + FRAME_PAD_X,
                MaxY = ys.Max() + NODE_HEIGHT + FRAME_PAD_BOTTOM
            };
        }

        public void ResolveCollisions(
            Dictionary<int, List<string>> vlanToNodes,
            Dictionary<string, (double X, double Y)> positions,
            int maxIterations = 30)
        {
            var vlans = vlanToNodes.Keys.OrderBy(v => v).ToList();
            if (vlans.Count <= 1) return;

            for (int iteration = 0; iteration < maxIterations; iteration++)
            {
                bool moved = false;

                // X座標の中心でソート
                var vlanBounds = new List<(double CenterX, int VLAN, Bounds Bounds)>();
                foreach (var vlan in vlans)
                {
                    var bounds = CalculateFrameBounds(vlanToNodes[vlan], positions);
                    if (bounds != null)
                    {
                        vlanBounds.Add((bounds.CenterX, vlan, bounds));
                    }
                }

                vlanBounds = vlanBounds.OrderBy(vb => vb.CenterX).ToList();

                // 隣接するVLAN枠の衝突をチェック
                for (int i = 0; i < vlanBounds.Count - 1; i++)
                {
                    var left = vlanBounds[i];
                    var right = vlanBounds[i + 1];

                    // 垂直方向の重なりチェック
                    if (!left.Bounds.OverlapsVertically(right.Bounds))
                        continue;

                    // 水平方向の押し出し量を計算
                    double overlap = (left.Bounds.MaxX + VLAN_FRAME_GAP) - right.Bounds.MinX;

                    if (overlap > 0)
                    {
                        // 右側のVLANのノードを全て移動
                        foreach (var nodeID in vlanToNodes[right.VLAN])
                        {
                            if (positions.ContainsKey(nodeID))
                            {
                                var (x, y) = positions[nodeID];
                                positions[nodeID] = (x + overlap, y);
                            }
                        }
                        moved = true;
                    }
                }

                if (!moved) break;
            }
        }
    }
}
