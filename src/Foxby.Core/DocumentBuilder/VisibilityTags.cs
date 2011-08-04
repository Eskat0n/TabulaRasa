using System.Collections.Generic;

namespace Foxby.Core.DocumentBuilder
{
    /// <summary>
    /// Контейнер для тегов
    /// </summary>
    public class VisibilityTags
    {
        public IEnumerable<string> NotUsingTagNames { get; private set; }
        public string UsingTagName { get; private set; }

        public VisibilityTags(string usingTagName, IEnumerable<string> notUsingTagName)
        {
            UsingTagName = usingTagName;
            NotUsingTagNames = notUsingTagName;
        }
    }
}