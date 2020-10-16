using System.Threading.Tasks;

namespace PowerPointToPlaces
{
    interface ISendToPlace
    {
        bool CanSend(string note);
        Task SendAsync(string note);
    }
}
