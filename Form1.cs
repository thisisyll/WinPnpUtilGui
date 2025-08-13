using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PnpUtilGui
{
    public partial class Form1 : Form
    {
        // Handles column sorting for the ListView
        private readonly ListViewColumnSorter lvwColumnSorter;

        // Stores all retrieved drivers for filtering without re-running pnputil
        private List<DriverInfo> allDrivers;

        /// <summary>
        /// Maps pnputil output property names to English and Chinese for parsing robustness.
        /// This allows the parser to work correctly on systems with different display languages.
        /// </summary>
        private static readonly Dictionary<string, (string English, string Chinese)> DriverPropertyMap =
             new Dictionary<string, (string English, string Chinese)>
             {
                { "PublishedName", ("Published name", "發佈名稱") },
                 { "OriginalName", ("Original name", "原始名稱") },
                { "Provider", ("Provider name", "提供者名稱") },
                { "Class", ("Class name", "類別名稱") },
                { "DriverVersion", ("Driver Version", "驅動程式版本") }
             };

        public Form1()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
        }

        /// <summary>
        /// Executes pnputil.exe with the given arguments and returns output, error, and exit code.
        /// </summary>
        private (string output, string error, int exitCode) RunPnputil(string arguments)
        {
            using (var p = new Process())
            {
                p.StartInfo.FileName = Path.Combine(Environment.SystemDirectory, "pnputil.exe");
                p.StartInfo.Arguments = arguments;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;

                p.Start();
                string output = p.StandardOutput.ReadToEnd();
                string error = p.StandardError.ReadToEnd();
                p.WaitForExit();

                return (output, error, p.ExitCode);
            }
        }

        /// <summary>
        /// Converts pnputil text output into a structured list of DriverInfo objects.
        /// </summary>
        private List<DriverInfo> ParsePnpUtilOutput(string output)
        {
            var drivers = new List<DriverInfo>();
            var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            DriverInfo currentDriver = null;

            foreach (var line in lines)
            {
                // A new driver entry always starts with "Published name".
                if (line.StartsWith(DriverPropertyMap["PublishedName"].English) ||
                    line.StartsWith(DriverPropertyMap["PublishedName"].Chinese))
                {
                    // If we were already building a driver object, add it to the list before starting a new one.
                    if (currentDriver != null)
                    {
                        drivers.Add(currentDriver);
                    }
                    currentDriver = new DriverInfo
                    {
                        PublishedName = line.Split(':')[1].Trim()
                    };
                }
                else if (currentDriver != null)
                {
                    foreach (var entry in DriverPropertyMap)
                    {
                        if (entry.Key == "PublishedName") continue; // Already handled above.

                        if (line.Contains(entry.Value.English) || line.Contains(entry.Value.Chinese))
                        {
                            string value = line.Split(new[] { ':' }, 2)[1].Trim(); // Split only on the first colon to handle values that might contain colons.

                            // Assign the value to the correct property on the DriverInfo object.
                            switch (entry.Key)
                            {
                                case "OriginalName":
                                    currentDriver.OriginalName = value;
                                    break;
                                case "Provider":
                                    currentDriver.Provider = value;
                                    break;
                                case "DriverVersion":
                                    currentDriver.DriverVersion = value;
                                    break;
                                case "Class":
                                    currentDriver.Class = value;
                                    break;
                            }
                            break;
                        }
                    }
                }
            }
            // Add the last driver object to the list.
            if (currentDriver != null)
            {
                drivers.Add(currentDriver);
            }
            return drivers;
        }

        private void SetSearchPlaceholder(bool isActive)
        {
            searchTextBox.Text = isActive ? "" : "Search by Provider...";
            searchTextBox.ForeColor = isActive
                ? System.Drawing.SystemColors.ControlText
                : System.Drawing.SystemColors.GrayText;
        }

        /// <summary>
        /// Clears and populates the ListView with a given list of drivers.
        /// </summary>
        private void PopulateListView(List<DriverInfo> driversToDisplay)
        {
            var items = driversToDisplay.Select(driver =>
            {
                var item = new ListViewItem(driver.PublishedName);
                item.SubItems.Add(driver.OriginalName);
                item.SubItems.Add(driver.Provider);
                item.SubItems.Add(driver.DriverVersion);
                item.SubItems.Add(driver.Class);
                return item;
            }).ToArray();

            listView1.Items.Clear();
            listView1.Items.AddRange(items);
        }

        private void SortListView() => listView1.Sort();

        /// <summary>
        /// Executes 'pnputil /enum-drivers', parses the output, and populates the ListView.
        /// </summary>
        private void EnumerateDriversButton_Click(object sender, EventArgs e)
        {
            var (output, _, _) = RunPnputil("/enum-drivers");
            allDrivers = ParsePnpUtilOutput(output);

            lvwColumnSorter.SortColumn = 2; // Provider name
            lvwColumnSorter.Order = SortOrder.Ascending;

            ApplyFilter(searchTextBox.Text);
        }

        /// <summary>
        /// Handles the column click event to update sorting order and column.
        /// </summary>
        private void ListView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // If the user clicks a column that is already being sorted, reverse the sort order.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                lvwColumnSorter.Order = (lvwColumnSorter.Order == SortOrder.Ascending) ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                // Otherwise, set the new column to sort by and default to ascending order.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }
            // Apply the new sorting options.
            SortListView();
        }

        /// <summary>
        /// Initiates filtering when the search button is clicked.
        /// </summary>
        private void SearchButton_Click(object sender, EventArgs e)
        {
            ApplyFilter(searchTextBox.Text);
        }

        /// <summary>
        /// Allows the user to press Enter in the search box to initiate a search.
        /// </summary>
        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ApplyFilter(searchTextBox.Text);
                e.Handled = true; // Prevent the 'ding' sound on Enter.
                e.SuppressKeyPress = true; // Stop the event from being passed to other controls.
            }
        }

        private void SearchTextBox_Enter(object sender, EventArgs e)
        {
            if (searchTextBox.Text == "Search by Provider...") SetSearchPlaceholder(true);
        }

        private void SearchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(searchTextBox.Text)) SetSearchPlaceholder(false);
        }

        /// <summary>
        /// Handles the logic for uninstalling one or more selected drivers.
        /// </summary>
        private void UninstallDriverToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one driver to uninstall.",
                    "No Driver Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var sb = new StringBuilder();
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                sb.AppendLine($"- {item.SubItems[0].Text} ({item.SubItems[1].Text})");
            }

            DialogResult confirmResult = MessageBox.Show(
                $"Are you sure you want to uninstall the following driver(s)?\n\n{sb}\n" +
                $"This operation requires administrator privileges and may require a system restart.",
                "Confirm Uninstall", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (confirmResult != DialogResult.OK) return;

            // Accumulate results
            StringBuilder successSb = new StringBuilder();
            StringBuilder failSb = new StringBuilder();

            foreach (ListViewItem item in listView1.SelectedItems)
            {
                string publishedName = item.SubItems[0].Text;
                var (output, error, exitCode) = RunPnputil($"/delete-driver {publishedName} /uninstall");

                if (exitCode == 0)
                    successSb.AppendLine($"{publishedName} - Success\nOutput:\n{output}\n");
                else
                    failSb.AppendLine($"{publishedName} - Failed\nError:\n{error}\nOutput:\n{output}\n");
            }

            // Show **one summary** dialog
            StringBuilder finalMessage = new StringBuilder();
            if (successSb.Length > 0) finalMessage.AppendLine("Successfully uninstalled drivers:\n" + successSb);
            if (failSb.Length > 0) finalMessage.AppendLine("Failed to uninstall drivers:\n" + failSb);

            if (finalMessage.Length > 0)
            {
                MessageBox.Show(finalMessage.ToString(), "Uninstall Summary", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            EnumerateDriversButton_Click(this, EventArgs.Empty);
        }

        /// <summary>
        /// Filters the cached list of allDrivers based on the search text and updates the ListView.
        /// </summary>
        private void ApplyFilter(string searchText)
        {
            if (allDrivers == null || !allDrivers.Any())
            {
                MessageBox.Show("Please enumerate drivers first by clicking 'Enumerate Drivers' button.",
                    "No Drivers Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                listView1.Items.Clear();
                return;
            }

            string query = searchText.Trim();
            bool isEmptySearch = string.IsNullOrEmpty(query) || query == "Search by Provider...";

            var filteredDrivers = allDrivers
                .Where(d => isEmptySearch || (d.Provider != null &&
                    d.Provider.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0))
                .ToList();

            PopulateListView(filteredDrivers);
            SortListView();
        }      
    }

    public class ListViewColumnSorter : IComparer
    {
        private int ColumnToSort;
        private SortOrder OrderOfSort;
        private readonly CaseInsensitiveComparer ObjectCompare;

        public ListViewColumnSorter()
        {
            ColumnToSort = 0;
            OrderOfSort = SortOrder.None;
            ObjectCompare = new CaseInsensitiveComparer();
        }

        /// <summary>
        /// Compares two ListViewItem objects for sorting.
        /// </summary>
        public int Compare(object x, object y)
        {
            var listviewX = (ListViewItem)x;
            var listviewY = (ListViewItem)y;

            int compareResult = ObjectCompare.Compare(
                listviewX.SubItems[ColumnToSort].Text,
                listviewY.SubItems[ColumnToSort].Text);

            // Adjust the result based on the desired sort order.
            if (OrderOfSort == SortOrder.Ascending)
            {
                return compareResult;
            }
            else if (OrderOfSort == SortOrder.Descending)
            {
                return -compareResult;
            }
            else
            {
                return 0; // No sorting.
            }
        }

        public int SortColumn
        {
            set => ColumnToSort = value;
            get => ColumnToSort;
        }

        public SortOrder Order
        {
            set => OrderOfSort = value;
            get => OrderOfSort;
        }
    }

    /// <summary>
    /// A data class to hold information about a single driver package.
    /// </summary>
    public class DriverInfo
    {
        public string PublishedName { get; set; }
        public string OriginalName { get; set; }
        public string Provider { get; set; }
        public string Class { get; set; }
        public string DriverVersion { get; set; }
    }
}
