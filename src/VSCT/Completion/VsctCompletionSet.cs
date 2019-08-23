using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MadsKristensen.ExtensibilityTools.Vsct
{
    class VsctCompletionSet : CompletionSet
    {
        private readonly List<Completion> _allCompletions;

        public VsctCompletionSet(string moniker, string displayName, ITrackingSpan applicableTo, List<Completion> completions, IEnumerable<Completion> completionBuilders)
            : base(moniker, displayName, applicableTo, completions, completionBuilders)
        {
            _allCompletions = completions;
        }

        public override void SelectBestMatch()
        {
            ITextSnapshot snapshot = ApplicableTo.TextBuffer.CurrentSnapshot;
            string typedText = ApplicableTo.GetText(snapshot);

            if (string.IsNullOrWhiteSpace(typedText))
            {
                if (this.WritableCompletions.Any())
                    SelectionStatus = new CompletionSelectionStatus(WritableCompletions.First(), true, true);

                return;
            }

            foreach (Completion comp in WritableCompletions)
            {
                if (comp.DisplayText.IndexOf(typedText, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    SelectionStatus = new CompletionSelectionStatus(comp, true, true);
                    return;
                }
            }
        }

        public override void Filter()
        {
            ITextSnapshot snapshot = ApplicableTo.TextBuffer.CurrentSnapshot;
            string typedText = ApplicableTo.GetText(snapshot);

            if (!string.IsNullOrEmpty(typedText))
            {
                List<Completion> temp = _allCompletions.Where(c => c.DisplayText.IndexOf(typedText, StringComparison.InvariantCultureIgnoreCase) > -1).ToList();

                if (temp.Any())
                {
                    this.WritableCompletions.BeginBulkOperation();

                    this.WritableCompletions.Clear();

                    this.WritableCompletions.AddRange(temp);

                    this.WritableCompletions.EndBulkOperation();
                }
            }
            else
            {
                this.WritableCompletions.BeginBulkOperation();

                this.WritableCompletions.Clear();

                this.WritableCompletions.AddRange(_allCompletions);

                this.WritableCompletions.EndBulkOperation();
            }
        }
    }
}
