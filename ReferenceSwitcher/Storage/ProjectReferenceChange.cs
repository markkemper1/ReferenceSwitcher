using System;

namespace ReferenceSwitcher
{
    [Serializable]
    public class ProjectReferenceChange
    {
        public string SourceProject { get; set; }

        public string ReferencedProject { get; set; }

        public string KnownPaths { get; set; }
    }
}