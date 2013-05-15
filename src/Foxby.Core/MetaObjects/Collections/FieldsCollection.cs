﻿namespace TabulaRasa.MetaObjects.Collections
{
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
	/// Collection which contains all sdt fields from document represented as <see cref="Field"/>s
	/// </summary>
	public class FieldsCollection : IEnumerable<Field>
	{
		private readonly IEnumerable<SdtElement> _elements;

		internal FieldsCollection(IEnumerable<SdtElement> elements)
		{
			_elements = elements;
		}

		/// <summary>
		/// Looks for sdt field in document specified by <paramref name="name"/>, <paramref name="tag"/> or both <paramref name="name"/> and <paramref name="tag"/>
		/// </summary>
		/// <param name="name">Name of field</param>
		/// <param name="tag">Name of tag of field</param>
		public bool Contains(string name = null, string tag = null)
		{
			return Fields.Any(x => SearchPredicate(x, name, tag));
		}

		public IEnumerator<Field> GetEnumerator()
		{
			return Fields.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}

		private IEnumerable<Field> Fields
		{
			get { return _elements.Select(x => new Field(x)); }
		}

		private static bool SearchPredicate(Field field, string name, string tag)
		{
			var fieldName = field.Name;
			var fieldTag = field.Tag;

			if (name != null && tag != null)
				return fieldName == name && fieldTag == tag;
			if (name != null)
				return fieldName == name;
			if (tag != null)
				return fieldTag == tag;
			return true;
		}
	}
}