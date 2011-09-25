using System;
using System.Collections.Generic;
using System.Linq;

namespace ReferenceSwitcher
{
    [Serializable]
    public class ProjectReferenceState
    {
        public ProjectReferenceChange[] Changes { get; set; }

        public void AddUpdate(string sourceProject, string referenceProject, string knownPath)
        {
            if(this.Changes == null) this.Changes = new ProjectReferenceChange[0];

            var item = this.Changes.Where(x => x.SourceProject == sourceProject && x.ReferencedProject == referenceProject).FirstOrDefault();

            if (item == null)
            {
                item = new ProjectReferenceChange()
                           {
                               SourceProject = sourceProject,
                               ReferencedProject = referenceProject
                           };
                Changes = Changes.Union(new[] {item}).ToArray();
            }

            item.KnownPaths = knownPath;
        }

        public IEnumerable<ProjectReferenceChange> GetProjectsThatReference(string uniqueName)
        {
            if (this.Changes == null) yield break;

            foreach (var item in this.Changes.Where(x => x.ReferencedProject == uniqueName))
            {
                yield return item;
            }
        }
    }
}