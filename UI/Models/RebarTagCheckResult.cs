using System.Collections.Generic;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarTagCheckResult
    {
        public IReadOnlyList<string> Structure { get; }
        public IReadOnlyList<string> Bending { get; }

        public RebarTagCheckResult(IReadOnlyList<string> structure, IReadOnlyList<string> bending)
        {
            Structure = structure ?? new List<string>();
            Bending = bending ?? new List<string>();
        }
    }
}

