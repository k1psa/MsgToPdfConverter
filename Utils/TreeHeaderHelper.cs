using System.Collections.Generic;

namespace MsgToPdfConverter.Utils
{
    public static class TreeHeaderHelper
    {
        // Builds a tree-like header from a parent chain and item label
        public static string BuildTreeHeader(List<string> parentChain, string itemLabel)
        {
            if (parentChain == null || parentChain.Count == 0)
                return itemLabel;
            var lines = new List<string>();
            for (int i = 0; i < parentChain.Count; i++)
            {
                string prefix = "";
                if (i > 0)
                    prefix = new string(' ', (i - 1) * 3) + (i == parentChain.Count - 1 ? "└── " : "├── ");
                lines.Add(prefix + parentChain[i]);
            }
            // Indent the item label to match the last parent
            string itemPrefix = new string(' ', parentChain.Count * 3) + "└── ";
            lines.Add(itemPrefix + itemLabel);
            return string.Join("\n", lines);
        }
    }
}
