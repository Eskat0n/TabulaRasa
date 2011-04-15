namespace Foxby.Core.MetaObjects
{
	public abstract class Format
	{
		public abstract void Invoke(IDocumentContextFormattableBuilder builder);

		public static implicit operator Format(string value)
		{
			return new NoneFormat(value);
		}

		public static Format operator +(string x, Format y)
		{
			return Concat(x, y);
		}

		public static Format operator +(Format x, Format y)
		{
			return Concat(x, y);
		}

		private static Format Concat(Format x, Format y)
		{
			return new ConcatFormat(x, y);
		}

		public Format Bold()
		{
			return new BoldFormat(this);
		}

		public Format Underlined()
		{
			return new UnderlinedFormat(this);
		}

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

			public override void Invoke(IDocumentContextFormattableBuilder builder)
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

			public override void Invoke(IDocumentContextFormattableBuilder builder)
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

			public override void Invoke(IDocumentContextFormattableBuilder builder)
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

			public override void Invoke(IDocumentContextFormattableBuilder builder)
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

			public override void Invoke(IDocumentContextFormattableBuilder builder)
			{
				format.Invoke(builder.Underlined);
			}
		}

		#endregion
	}
}