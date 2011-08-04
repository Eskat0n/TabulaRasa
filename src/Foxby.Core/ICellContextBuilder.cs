using System;

namespace Foxby.Core
{
	/// <summary>
	/// Provides methods for building content of table cells
	/// </summary>
	public interface ICellContextBuilder : IDocumentContextBuilder, IDocumentFormattableBuilder<ICellContextBuilder>
	{
		/// <summary>
		/// Sets text <paramref name="content"/> in cell
		/// </summary>
		/// <param name="content">Text content for cell</param>
		void Cell(string content);

		/// <summary>
		/// Builds cell <paramref name="content"/> using delegate
		/// </summary>
		/// <param name="content">Content builder</param>
	    void Cell(Action<IDocumentContextBuilder> content);
	}
}