using System.Collections.Generic;
using System.Linq;

namespace NetworkDiagramApp
{
    public class TreeBuilder
    {
        public Dictionary<string, NetworkNode> Nodes { get; }
        public Dictionary<string, List<string>> Children { get; }
        public Dictionary<string, string> PrimaryParent { get; }
        public List<string> Roots { get; }

        public TreeBuilder(Dictionary<string, NetworkNode> nodes)
        {
            Nodes = nodes;
            Children = new Dictionary<string, List<string>>();
            PrimaryParent = new Dictionary<string, string>();

            // 初期化
            foreach (var nodeID in nodes.Keys)
            {
                Children[nodeID] = new List<string>();
            }

            // 親子関係を構築
            foreach (var kvp in nodes)
            {
                string childID = kvp.Key;
                var node = kvp.Value;

                if (node.Parents.Count == 0) continue;

                // 配置用の親（先頭のみ使用）
                string primaryParent = node.Parents[0];
                PrimaryParent[childID] = primaryParent;
                Children[primaryParent].Add(childID);
            }

            // 子をソート
            foreach (var key in Children.Keys.ToList())
            {
                Children[key] = Children[key].OrderBy(c => c).ToList();
            }

            // ルートノードを取得
            Roots = FindRoots();
        }

        private List<string> FindRoots()
        {
            var roots = Nodes.Keys.Where(nid => !PrimaryParent.ContainsKey(nid)).ToList();

            // ONUを先頭に
            var onuRoots = roots.Where(r => Nodes[r].IsONU()).OrderBy(r => r).ToList();
            var otherRoots = roots.Where(r => !Nodes[r].IsONU()).OrderBy(r => r).ToList();

            return onuRoots.Concat(otherRoots).ToList();
        }
    }
}
