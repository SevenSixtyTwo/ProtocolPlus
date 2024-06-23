using protocolPlus.Core;
using protocolPlus.MVVM.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

using static protocolPlus.Core.DatabaseUtils;
using static WordUtils;

namespace protocolPlus.MVVM.ViewModel
{
    class DiscoveryViewModel
    {
        public ObservableCollection<Tool> Tools { get; set; }
        public ObservableCollection<Tool> AvailableTools { get; set; }
        public ObservableCollection<DropDownItem> DropDownListItems { get; set; }
        public ObservableCollection<DataItem> DataGridItems { get; set; }

        public string templateFilePath = $@"template-std.docx";
        public string newFileName = "протокол ПСИ СТД";
        public string newFilePath = Environment.CurrentDirectory + @"\протоколы";

        public SQLiteConnection dbConnection = CreateConnection();
        public RelayCommand SaveProtocolCommand { get; set; }
        public Tool SelectedTool { get; set; }

        public ICommand AddToolCommand { get; private set; }
        public ICommand DeleteToolCommand { get; private set; }

        private DropDownItem _selectedItem;
        public DropDownItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                UpdateDataGridItems(); // Update data grid when selection changes
                OnPropertyChanged(nameof(SelectedItem));
            }
        }
        private void UpdateDataGridItems()
        {
            DataGridItems.Clear();
            if (SelectedItem != null)
            {
                foreach (var item in SelectedItem.Data)
                {
                    DataGridItems.Add(item);
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public int GetSelectedItemId()
        {
            if (SelectedItem != null)
                return SelectedItem.Id;
            else
                return -1;
        }



        private void AddTool()
        {
            if (SelectedTool != null)
            {
                Tools.Add(new Tool() // Create a copy of the selected tool
                {
                    Identifier = SelectedTool.Identifier,
                    Name = SelectedTool.Name,
                    Type = SelectedTool.Type,
                    AssurancNum = SelectedTool.AssurancNum,
                    VerificationNum = SelectedTool.VerificationNum,
                    VerificationDate = SelectedTool.VerificationDate

                });
                OnPropertyChanged(nameof(Tools));
            }
        }

        private void DeleteTool()
        {
            if (SelectedTool != null)
            {
                Tools.Remove(SelectedTool);
                SelectedTool = null; // Reset SelectedTool after deletion
                OnPropertyChanged(nameof(Tools));
                OnPropertyChanged(nameof(SelectedTool)); // Notify UI about selection change
            }
        }

        private bool CanDeleteTool()
        {
            return Tools.Count > 0;
        }

        public DiscoveryViewModel()
        {
            Tools = [];
            AvailableTools = [];

            DropDownListItems = [];
            DataGridItems = [];

            SelectedItem = DropDownListItems.FirstOrDefault();


            dbConnection.Open();

            var cmdTools = dbConnection.CreateCommand();
            cmdTools.CommandText = "SELECT id, name, type, assurance_num, verification_num, verification_date from tool;";

            var readerTools = cmdTools.ExecuteReader();
            while (readerTools.Read())
            {
                AvailableTools.Add(new Tool()
                {
                    Identifier = readerTools.GetInt32(0).ToString(),
                    Name = readerTools.GetString(1),
                    Type = readerTools.GetString(2),
                    AssurancNum = readerTools.GetString(3),
                    VerificationNum = readerTools.GetString(4),
                    VerificationDate = readerTools.GetString(5)
                });
            }

            AddToolCommand = new RelayCommand(o =>
            {
                AddTool();
            });
            DeleteToolCommand = new RelayCommand(o =>
            {
                if (CanDeleteTool())
                    DeleteTool();
            });

            SaveProtocolCommand = new RelayCommand(o =>
            {
                string currentDate = DateTime.Now.ToString("dd.MM.yyyy");
                string finalResult = "пригоден";

                Dictionary<string, string> tagsAndValues = [];

                tagsAndValues.Add("<protocol.creation_date>", currentDate);

                var newFullPathFile = CreateNewDocumentFromTemplate(templateFilePath, newFilePath, newFileName);

                int selectedMachineId = GetSelectedItemId();

                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineNameTag, DropDownListItems[selectedMachineId].Data[0].MachineName);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineAssuranceNumTag, DropDownListItems[selectedMachineId].Data[0].MachineAssuranceNum);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachinePowerTag, DropDownListItems[selectedMachineId].Data[0].MachinePower);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineVoltageStTag, DropDownListItems[selectedMachineId].Data[0].MachineVoltageSt);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCurrentStTag, DropDownListItems[selectedMachineId].Data[0].MachineCurrentSt);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineFrequencyTag, DropDownListItems[selectedMachineId].Data[0].MachineFrequency);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineRpmTag, DropDownListItems[selectedMachineId].Data[0].MachineRpm);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCosinusTag, DropDownListItems[selectedMachineId].Data[0].MachineCosinus);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineEfficencyTag, DropDownListItems[selectedMachineId].Data[0].MachineEfficency);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCurrentExcTag, DropDownListItems[selectedMachineId].Data[0].MachineCurrentExc);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineVoltageExcTag, DropDownListItems[selectedMachineId].Data[0].MachineVoltageExc);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineRotatioTag, DropDownListItems[selectedMachineId].Data[0].MachineRotatio);

                tagsAndValues.Add("<final_result>", finalResult);

                foreach (var item in tagsAndValues)
                {
                    ReplaceTagsInFile(newFullPathFile, item.Key, item.Value);
                }

                Dictionary<string, string> toolsTagsAndValues = [];

                string tableCaption = "TOOLS_TABLE";
                var rowPattern = GetRowPattern(newFullPathFile, tableCaption);
                foreach (var tool in Tools)
                {
                    toolsTagsAndValues.Add(tool.NameTag, tool.Name);
                    toolsTagsAndValues.Add(tool.TypeTag, tool.Type);
                    toolsTagsAndValues.Add(tool.AssuranceNumTag, tool.AssurancNum);
                    toolsTagsAndValues.Add(tool.VerificationNumTag, tool.VerificationNum);
                    toolsTagsAndValues.Add(tool.VerificationDateTag, tool.VerificationDate);
                    toolsTagsAndValues.Add(tool.ReadyTag, tool.Ready);

                    CreateRowWithPattern(newFullPathFile, tableCaption, rowPattern, toolsTagsAndValues);
                    toolsTagsAndValues.Clear();
                }
                DeleteRowPattern(newFullPathFile, tableCaption, rowPattern);
            });
        }
    }
}
