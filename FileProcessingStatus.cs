using System.ComponentModel;

namespace MsgToPdfConverter
{
    public class FileProcessingStatus : INotifyPropertyChanged
    {
        private string _fileName;
        private bool _isProcessing;
        private bool _isDone;

        public string FileName
        {
            get => _fileName;
            set { _fileName = value; OnPropertyChanged(nameof(FileName)); }
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set { _isProcessing = value; OnPropertyChanged(nameof(IsProcessing)); }
        }

        public bool IsDone
        {
            get => _isDone;
            set { _isDone = value; OnPropertyChanged(nameof(IsDone)); }
        }

        public FileProcessingStatus(string fileName)
        {
            FileName = fileName;
            IsProcessing = false;
            IsDone = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
