namespace TabulaRasa.MetaObjects
{
	///<summary>
	/// Stores formatting options for text strings
	///</summary>
	public abstract class Format
	{
		internal abstract void Invoke(IDocumentContextFormattableBuilder builder);

		///<summary>
		/// Converts string to empty formatted string
		///</summary>
		///<param name="value">String of text</param>
		public static implicit operator Format(string value)
		{
			return new NoneFormat(value);
		}

		///<summary>
		/// Concates formatted string with unformatted string
		///</summary>
		///<param name="x">String of text</param>
		///<param name="y">Formatted string</param>
		public static Format operator +(string x, Format y)
		{
			return Concat(x, y);
		}

		///<summary>
		/// Concates two formatted strings
		///</summary>
		///<param name="x">Left-side formatted string</param>
		///<param name="y">Right-side formatted string</param>
		public static Format operator +(Format x, Format y)
		{
			return Concat(x, y);
		}

		private static Format Concat(Format x, Format y)
		{
			return new ConcatFormat(x, y);
		}

		///<summary>
		/// Appends bold formatting option to current format
		///</summary>
		public Format Bold()
		{
			return new BoldFormat(this);
		}

		///<summary>
		/// Appends underlined formatting option to current format
		///</summary>
		public Format Underlined()
		{
			return new UnderlinedFormat(this);
		}

		///<summary>
		/// Appends italic formatting option to current format
		///</summary>
		public Format Italic()
		{
			return new ItalicFormat(this);
		}

		#region Nested type: BoldFormat

		private class BoldFormat : Format
		{
			private readonly Format format;

			public BoldFormat(Format format)
			{
				this.format = format;
			}

			internal override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				format.Invoke(builder.Bold);
			}
		}

		#endregion

		#region Nested type: ConcatFormat

		private class ConcatFormat : Format
		{
			private readonly Format x;
			private readonly Format y;

			public ConcatFormat(Format x, Format y)
			{
				this.x = x;
				this.y = y;
			}

			internal override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				x.Invoke(builder);
				y.Invoke(builder);
			}
		}

		#endregion

		#region Nested type: ItalicFormat

		private class ItalicFormat : Format
		{
			private readonly Format format;

			public ItalicFormat(Format format)
			{
				this.format = format;
			}

			internal override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				format.Invoke(builder.Italic);
			}
		}

		#endregion

		#region Nested type: NoneFormat

		private class NoneFormat : Format
		{
			private readonly string[] strings;

			public NoneFormat(params string[] strings)
			{
				this.strings = strings;
			}

			internal override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				builder.Text(strings);
			}
		}

		#endregion

		#region Nested type: UnderlinedFormat

		private class UnderlinedFormat : Format
		{
			private readonly Format format;

			public UnderlinedFormat(Format format)
			{
				this.format = format;
			}

			internal override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				format.Invoke(builder.Underlined);
			}
		}

		#endregion
	}
}