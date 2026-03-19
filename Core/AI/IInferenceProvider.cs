using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    public interface IInferenceProvider
    {
        Task<AIResult> GenerateFillAsync(AIContext context, CancellationToken cancellationToken);
    }
}

