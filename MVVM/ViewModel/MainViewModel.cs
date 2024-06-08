using protocolPlus.Core;

namespace protocolPlus.MVVM.ViewModel
{
    class MainViewModel : ObservableObject
    {
        public RelayCommand HomeViewCommand { get; set; }
        public RelayCommand DiscoveryViewCommand { get; set; }

        public ProtocolViewModel ProtocolVM { get; set; }
        public DiscoveryViewModel DiscoveryVM { get; set; }

        private object _currentView;
        
        public object CurrentView
        {
            get { return _currentView; }
            set
            {
                _currentView = value;
                OnPropetyChanged();
            }
        }
        public MainViewModel()
        {
            ProtocolVM = new ProtocolViewModel();
            DiscoveryVM = new DiscoveryViewModel();
            
            CurrentView = ProtocolVM;

            HomeViewCommand = new RelayCommand(o =>
            {
                CurrentView = ProtocolVM;
            });

            DiscoveryViewCommand = new RelayCommand(o =>
            {
                CurrentView = DiscoveryVM;
            });
        }
    }
}
