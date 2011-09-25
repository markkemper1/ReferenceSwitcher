using System;
using System.IO;
using System.Xml.Serialization;
using EnvDTE;

namespace ReferenceSwitcher
{
    public class StorageProvider
    {
        public void Save(Solution solution, ProjectReferenceState state)
        {
            var serializer = new XmlSerializer(typeof (ProjectReferenceState));
            using (var stream = File.OpenWrite(GetFilename(solution)))
            {
                serializer.Serialize(stream, state);
            }
        }

        public ProjectReferenceState Load(Solution solution)
        {
            string filename = GetFilename(solution);
            
            if(!File.Exists(filename)) return new ProjectReferenceState();

            var serializer = new XmlSerializer(typeof(ProjectReferenceState));
            using (var stream = File.OpenRead(filename))
            {
                return serializer.Deserialize(stream) as ProjectReferenceState;
            }
        }

        private string GetFilename(Solution solution)
        {
            string path = Path.GetDirectoryName(solution.FileName);
            return Path.Combine(path, Path.GetFileName(solution.FileName) + ".switchReferences.xml");
        }
    }
}
