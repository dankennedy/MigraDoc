#region MigraDoc - Creating Documents on the Fly

//
// Authors:
//   Klaus Potzesny (mailto:Klaus.Potzesny@pdfsharp.com)
//
// Copyright (c) 2001-2009 empira Software GmbH, Cologne (Germany)
//
// http://www.pdfsharp.com
// http://www.migradoc.com
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.

#endregion

using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Visitors;
using MigraDoc.Rendering.Resources;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace MigraDoc.Rendering
{
    /// <summary>
    ///     Provides the functionality to convert a MigraDoc document into PDF.
    /// </summary>
    public class PdfDocumentRenderer
    {
        /// <summary>
        ///     Initializes a new instance of the PdfDocumentRenderer class.
        /// </summary>
        public PdfDocumentRenderer()
        {
        }

        /// <summary>
        ///     Initializes a new instance of the PdfDocumentRenderer class.
        /// </summary>
        /// <param name="unicode">If true Unicode encoding is used for all text. If false, WinAnsi encoding is used.</param>
        public PdfDocumentRenderer(bool unicode)
        {
            Unicode = unicode;
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="PdfDocumentRenderer" /> class.
        /// </summary>
        /// <param name="unicode">If true Unicode encoding is used for all text. If false, WinAnsi encoding is used.</param>
        /// <param name="fontEmbedding">Specifies which fonts will be imbedded in the PDF file.</param>
        public PdfDocumentRenderer(bool unicode, PdfFontEmbedding fontEmbedding)
        {
            Unicode = unicode;
            FontEmbedding = fontEmbedding;
        }

        /// <summary>
        ///     Gets a value indicating whether the text is rendered as Unicode.
        /// </summary>
        public bool Unicode { get; }

        /// <summary>
        ///     Gets a value indicating whether all used fonts are embedded in the document.
        /// </summary>
        public PdfFontEmbedding FontEmbedding { get; }

        /// <summary>
        ///     Gets or sets the language.
        /// </summary>
        /// <value>The language.</value>
        public string Language { get; set; } = string.Empty;

        /// <summary>
        ///     Set the MigraDoc document to be rendered by this printer.
        /// </summary>
        public Document Document
        {
            set
            {
                document = null;
                value.BindToRenderer(this);
                document = value;
            }
        }

        Document document;

        /// <summary>
        ///     Gets or sets a document renderer.
        /// </summary>
        /// <remarks>
        ///     A document renderer is automatically created and prepared
        ///     when printing before this property was set.
        /// </remarks>
        public DocumentRenderer DocumentRenderer
        {
            get
            {
                if (documentRenderer == null)
                    PrepareDocumentRenderer();
                return documentRenderer;
            }
            set { documentRenderer = value; }
        }

        DocumentRenderer documentRenderer;

        void PrepareDocumentRenderer()
        {
            PrepareDocumentRenderer(false);
        }

        void PrepareDocumentRenderer(bool prepareCompletely)
        {
            if (document == null)
                throw new InvalidOperationException(Messages.PropertyNotSetBefore("DocumentRenderer",
                    MethodBase.GetCurrentMethod().Name));

            if (documentRenderer == null)
            {
                documentRenderer = new DocumentRenderer(document);
                documentRenderer.WorkingDirectory = WorkingDirectory;
            }
            if (prepareCompletely && documentRenderer.formattedDocument == null)
            {
                documentRenderer.PrepareDocument();
            }
        }

        /// <summary>
        ///     Renders the document into a PdfDocument containing all pages of the document.
        /// </summary>
        public void RenderDocument()
        {
#if true
            PrepareRenderPages();
#else
      if (this.documentRenderer == null)
        PrepareDocumentRenderer();

      if (this.pdfDocument == null)
      {
        this.pdfDocument = new PdfDocument();
        this.pdfDocument.Info.Creator = VersionInfo.Creator;
      }

      WriteDocumentInformation();
#endif
            RenderPages(1, documentRenderer.FormattedDocument.PageCount);
        }

        /// <summary>
        ///     Renders the document into a PdfDocument containing all pages of the document.
        /// </summary>
        public void PrepareRenderPages()
        {
            //if (this.documentRenderer == null)
            PrepareDocumentRenderer(true);

            if (PdfDocument == null)
            {
                PdfDocument = CreatePdfDocument();
                if (document.UseCmykColor)
                    PdfDocument.Options.ColorMode = PdfColorMode.Cmyk;
            }

            WriteDocumentInformation();
            //RenderPages(1, this.documentRenderer.FormattedDocument.PageCount);
        }

        /// <summary>
        ///     Gets the count of pages.
        /// </summary>
        public int PageCount
        {
            get { return documentRenderer.FormattedDocument.PageCount; }
        }

        /// <summary>
        ///     Saves the PdfDocument to the specified path. If a file already exists, it will be overwritten.
        /// </summary>
        public void Save(string path)
        {
            if (path == null)
                throw new ArgumentNullException("path");

            if (path == "")
                throw new ArgumentException("PDF file Path must not be empty");

            if (WorkingDirectory != null)
                Path.Combine(WorkingDirectory, path);

            PdfDocument.Save(path);
        }

        /// <summary>
        ///     Saves the PDF document to the specified stream.
        /// </summary>
        public void Save(Stream stream, bool closeStream)
        {
            PdfDocument.Save(stream, closeStream);
        }

        /// <summary>
        ///     Renders the spcified page range.
        /// </summary>
        /// <param name="startPage">The first page to print.</param>
        /// <param name="endPage">The last page to print</param>
        public void RenderPages(int startPage, int endPage)
        {
            if (startPage < 1)
                throw new ArgumentOutOfRangeException("startPage");

            if (endPage > documentRenderer.FormattedDocument.PageCount)
                throw new ArgumentOutOfRangeException("endPage");

            if (documentRenderer == null)
                PrepareDocumentRenderer();

            if (PdfDocument == null)
                PdfDocument = CreatePdfDocument();

            documentRenderer.printDate = DateTime.Now;
            for (var pageNr = startPage; pageNr <= endPage; ++pageNr)
            {
                var pdfPage = PdfDocument.AddPage();
                var pageInfo = documentRenderer.FormattedDocument.GetPageInfo(pageNr);
                pdfPage.Width = pageInfo.Width;
                pdfPage.Height = pageInfo.Height;
                pdfPage.Orientation = pageInfo.Orientation;

                using (var gfx = XGraphics.FromPdfPage(pdfPage))
                {
                    gfx.MUH = Unicode ? PdfFontEncoding.Unicode : PdfFontEncoding.WinAnsi;
                    gfx.MFEH = FontEmbedding;
                    documentRenderer.RenderPage(gfx, pageNr);
                }
            }
        }

        /// <summary>
        ///     Gets or sets a working directory for the printing process.
        /// </summary>
        public string WorkingDirectory { get; set; }

        /// <summary>
        ///     Gets or sets the PDF document to render on.
        /// </summary>
        /// <remarks>A PDF document in memory is automatically created when printing before this property was set.</remarks>
        public PdfDocument PdfDocument { get; set; }

        /// <summary>
        ///     Writes document information like author and subject to the PDF document.
        /// </summary>
        public void WriteDocumentInformation()
        {
            if (!document.IsNull("Info"))
            {
                var docInfo = document.Info;
                var pdfInfo = PdfDocument.Info;

                if (!docInfo.IsNull("Author"))
                    pdfInfo.Author = docInfo.Author;

                if (!docInfo.IsNull("Keywords"))
                    pdfInfo.Keywords = docInfo.Keywords;

                if (!docInfo.IsNull("Subject"))
                    pdfInfo.Subject = docInfo.Subject;

                if (!docInfo.IsNull("Title"))
                    pdfInfo.Title = docInfo.Title;

                if (CustomProperties != null)
                {
                    foreach (var prop in CustomProperties)
                    {
                        pdfInfo.Elements.Add(!prop.Key.StartsWith("/")
                            ? new KeyValuePair<string, PdfItem>("/" + prop.Key, prop.Value)
                            : prop);
                    }
                }
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<KeyValuePair<string, PdfItem>> CustomProperties { get; set; }        

        /// <summary>
        ///     Creates a new PDF document.
        /// </summary>
        PdfDocument CreatePdfDocument()
        {
            var document = new PdfDocument();
            document.Info.Creator = VersionInfo.Creator;
            if (Language != null && Language.Length != 0)
                document.Language = Language;
            return document;
        }
    }
}