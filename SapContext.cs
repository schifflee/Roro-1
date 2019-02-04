

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Roro
{
    public sealed class SapContext : Context
    {
        private XObject appObject;

        public static readonly SapContext Shared = new SapContext();

        private SapContext()
        {
            this.IsAlive();
        }

        public override Element GetElementFromPoint(int screenX, int screenY) =>
            this.IsAlive()
            && this.appObject.Get("ActiveSession") is XObject session
            && session.Invoke("FindByPosition", screenX, screenY, false) is XObject rawElementInfo
            && rawElementInfo.Invoke<string>("Item", 0) is string rawElementId
            && session.Invoke("FindById", rawElementId, false) is XObject rawElement
            ? new SapElement(rawElement) : null;

        public override IEnumerable<Element> GetElementsFromQuery(Query query)
        {
            var result = new List<SapElement>();
            var candidates = new Queue<SapElement>();
            var targetPath = query.First(x => x.Name == "Path").Value.ToString();

            if (!this.IsAlive())
            {
                return result;
            }

            foreach (var connection in this.appObject.Get("Connections"))
            {
                foreach (var session in connection.Get("Sessions"))
                {
                    candidates.Enqueue(new SapElement(session));
                }
            }

            while (candidates.Count > 0)
            {
                var candidate = candidates.Dequeue();
                var candidatePath = candidate.Path;
                if (targetPath.StartsWith(candidatePath))
                {
                    if (targetPath.Equals(candidatePath))
                    {
                        if (candidate.TryQuery(query))
                        {
                            result.Add(candidate);
                        }
                    }
                    else
                    {
                        foreach (var child in candidate.Children)
                        {
                            candidates.Enqueue(child);
                        }
                    }
                }
            }
            return result;
        }


        private bool IsAlive()
        {
            if (this.ProcessId > 0 && Process.GetProcessById(this.ProcessId) is Process proc)
            {
                return true;
            }

            if (Type.GetTypeFromProgID("SapROTWr.SapROTWrapper") is Type type
                && new XObject(Activator.CreateInstance(type)) is XObject ROTWrapper
                && ROTWrapper.Invoke("GetROTEntry", "SAPGUI") is XObject ROTEntry
                && ROTEntry.Invoke("GetScriptingEngine") is XObject appObject
                && Process.GetProcessesByName("saplogon").FirstOrDefault() is Process saplogon)
            {
                this.appObject = appObject;
                this.ProcessId = saplogon.Id;
                return true;
            }
            else
            {
                this.appObject = null;
                this.ProcessId = 0;
                return false;
            }
        }
    }
}
