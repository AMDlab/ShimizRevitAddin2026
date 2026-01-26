using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Model
{
    internal class RebarTag
    {
        public ElementId RebarId { get; }
        public ElementId ViewId { get; }

        public IReadOnlyList<ElementId> MatchedTagIds { get; }
        public IReadOnlyList<ElementId> StructureTagIds { get; }
        public IReadOnlyList<ElementId> BendingDetailTagIds { get; }

        public RebarTag(
          ElementId rebarId,
          ElementId viewId,
          IReadOnlyList<ElementId> matchedTagIds,
          IReadOnlyList<ElementId> structureTagIds,
          IReadOnlyList<ElementId> bendingDetailTagIds)
        {
            RebarId = rebarId;
            ViewId = viewId;
            MatchedTagIds = matchedTagIds ?? new List<ElementId>();
            StructureTagIds = structureTagIds ?? new List<ElementId>();
            BendingDetailTagIds = bendingDetailTagIds ?? new List<ElementId>();
        }
    }
}
