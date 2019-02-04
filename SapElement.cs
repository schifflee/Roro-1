

using System;
using System.Collections.Generic;
using System.Drawing;

namespace Roro
{
    public sealed class SapElement : Element
    {
        private readonly XObject rawElement;

        internal SapElement(XObject rawElement)
        {
            this.rawElement = rawElement;
        }

        [BotProperty]
        public string Id => this.rawElement.Get<string>("Id");

        [BotProperty]
        public string Name => this.rawElement.Get<string>("Name");

        [BotProperty]
        public string Type => (this.rawElement.Get<SapElementType>("TypeAsNumber")).ToString().ToLower();

        [BotProperty]
        public override string Path => string.Format("{0}/{1}", this.Parent == null ? string.Empty : this.Parent.Path, this.Type);

        public override Rectangle Bounds => new Rectangle(
            this.rawElement.Get<int>("ScreenLeft"),
            this.rawElement.Get<int>("ScreenTop"),
            this.rawElement.Get<int>("Width"),
            this.rawElement.Get<int>("Height"));

        public SapElement Parent => this.rawElement.Get("Parent") is XObject rawParent ? new SapElement(rawParent) : null;

        public IEnumerable<SapElement> Children
        {
            get
            {
                var children = new List<SapElement>();
                var rawChildren = this.rawElement.Get("Children");
                foreach (var rawChild in rawChildren)
                {
                    children.Add(new SapElement(rawChild));
                }
                return children;
            }
        }
    }
}
