namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarTagPairRow
    {
        public string StructureRebarTag { get; }
        public string BendingDetailRebarTag { get; }

        public RebarTagPairRow(string structureRebarTag, string bendingDetailRebarTag)
        {
            StructureRebarTag = structureRebarTag ?? string.Empty;
            BendingDetailRebarTag = bendingDetailRebarTag ?? string.Empty;
        }
    }
}

