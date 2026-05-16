/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System.Xml.Linq;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Cached snapshot of the selection state on a OneNote page. Built once per
    /// <c>ShowForm</c> invocation by <see cref="From"/> and threaded into the
    /// helpers that previously each re-ran the same
    /// <c>Descendants(ns + "Outline").Where(selected == all|partial)</c>
    /// traversal (resolves review item 2.1 in
    /// <c>.local/docs/reviews/03-code-quality.md</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Two outline references are recorded because the legacy helpers used
    /// slightly different filters:
    /// </para>
    /// <list type="bullet">
    ///   <item>
    ///     <description>
    ///       <c>GetSelectedText</c>, <c>GetOutline</c> and
    ///       <c>IsSelectedTextInline</c> looked for any outline with
    ///       <c>selected="all"</c> OR <c>selected="partial"</c>; that match
    ///       lives in <see cref="Outline"/>.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       <c>GetMousePointPosition</c> looked specifically for
    ///       <c>selected="partial"</c> (the outline that owns the caret
    ///       <c>&lt;Position&gt;</c> element); that narrower match lives in
    ///       <see cref="PartialOutline"/>. In practice the two refer to the same
    ///       element when a partial selection exists, but keeping them distinct
    ///       preserves the pre-refactor behaviour bit-for-bit.
    ///     </description>
    ///   </item>
    /// </list>
    /// <para>
    /// <see cref="Table"/> is the all-or-partial <c>&lt;Table&gt;</c> nested
    /// inside <see cref="Outline"/>, or <c>null</c> when the selection is not
    /// inside a table cell. The legacy branch <c>table == null ? ... : ...</c>
    /// duplicated almost verbatim between <c>GetSelectedText</c> and
    /// <c>IsSelectedTextInline</c> reduces to a single <c>Table != null</c>
    /// check against this field.
    /// </para>
    /// </remarks>
    internal sealed class PageSelection
    {
        /// <summary>
        /// The first <c>&lt;Outline&gt;</c> on the page carrying
        /// <c>selected="all"</c> or <c>selected="partial"</c>, or <c>null</c>
        /// if no outline is selected.
        /// </summary>
        public XElement Outline { get; }

        /// <summary>
        /// The first <c>&lt;Outline&gt;</c> on the page carrying
        /// <c>selected="partial"</c> exclusively (the one that owns the caret
        /// <c>&lt;Position&gt;</c>), or <c>null</c>. Used by the legacy mouse
        /// point lookup.
        /// </summary>
        public XElement PartialOutline { get; }

        /// <summary>
        /// The first <c>&lt;Table&gt;</c> nested inside <see cref="Outline"/>
        /// carrying <c>selected="all"</c> or <c>selected="partial"</c>, or
        /// <c>null</c> if the selection is not inside a table cell.
        /// </summary>
        public XElement Table { get; }

        private PageSelection(XElement outline, XElement partialOutline, XElement table)
        {
            Outline = outline;
            PartialOutline = partialOutline;
            Table = table;
        }

        /// <summary>
        /// Single source of truth for selection detection. Walks
        /// <paramref name="pageRoot"/> once and resolves every state the four
        /// legacy helpers used to derive independently. Returns a non-null
        /// instance even when nothing is selected so callers can stay
        /// branch-free; the individual properties are then null.
        /// </summary>
        /// <param name="pageRoot">The page <c>&lt;Page&gt;</c> root, or
        /// <c>null</c> if the page XML could not be parsed.</param>
        /// <param name="ns">The OneNote XML namespace (passed in rather than
        /// read from <c>AddIn.ns</c> so the helper does not depend on
        /// COM-add-in state).</param>
        public static PageSelection From(XElement pageRoot, XNamespace ns)
        {
            if (pageRoot == null)
            {
                return new PageSelection(null, null, null);
            }

            XElement outline = null;
            XElement partialOutline = null;

            foreach (var candidate in pageRoot.Descendants(ns + "Outline"))
            {
                var selectedAttr = candidate.Attribute("selected");
                if (selectedAttr == null)
                {
                    continue;
                }
                var value = selectedAttr.Value;
                bool isAll = value == "all";
                bool isPartial = value == "partial";
                if (outline == null && (isAll || isPartial))
                {
                    outline = candidate;
                }
                if (partialOutline == null && isPartial)
                {
                    partialOutline = candidate;
                }
                if (outline != null && partialOutline != null)
                {
                    break;
                }
            }

            XElement table = null;
            if (outline != null)
            {
                foreach (var candidate in outline.Descendants(ns + "Table"))
                {
                    var selectedAttr = candidate.Attribute("selected");
                    if (selectedAttr == null)
                    {
                        continue;
                    }
                    var value = selectedAttr.Value;
                    if (value == "all" || value == "partial")
                    {
                        table = candidate;
                        break;
                    }
                }
            }

            return new PageSelection(outline, partialOutline, table);
        }
    }
}
