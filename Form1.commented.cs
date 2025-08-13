using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Collections;

namespace PnpUtilGui
{
    public partial class Form1 : Form
    {
        // Sorter class for handling ListView column sorting logic.
        private readonly ListViewColumnSorter lvwColumnSorter;
        // Caches all drivers retrieved from pnputil to allow for fast filtering without re-running the process.
        private List<DriverInfo> allDrivers;

        public Form1()
        {
            InitializeComponent();
            // Associate the custom sorter with the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
        }

        /// <summary>
        /// Executes 'pnputil /enum-drivers', parses the output, and populates the ListView.
        /// </summary>
        private void EnumerateDriversButton_Click(object sender, EventArgs e)
        {
            Process p = new Process();
            // Use the full path to pnputil.exe to avoid issues with system PATH environment variables.
            p.StartInfo.FileName = Path.Combine(Environment.SystemDirectory, "pnputil.exe");
            p.StartInfo.Arguments = "/enum-drivers";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            // Parse the raw text output into a list of DriverInfo objects.
            allDrivers = ParsePnpUtilOutput(output);

            // Set a default sort column and order. The initial population will be sorted by this.
            lvwColumnSorter.SortColumn = 2; // Default to "Provider name" column
            lvwColumnSorter.Order = SortOrder.Ascending;

            // Apply the initial filter (which is empty, showing all drivers) and sort the view.
            ApplyFilter(searchTextBox.Text);
        }

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

        /// <summary>
        /// Parses the multi-line string output from 'pnputil /enum-drivers' into a list of DriverInfo objects.
        /// </summary>
        /// <param name="output">The raw string output from the pnputil process.</param>
        /// <returns>A list of structured DriverInfo objects.</returns>
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
                    // This line is a property of the current driver object.
                    // Iterate through our known properties to find a match.
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
                            break; // Property found and assigned, move to the next line.
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

        /// <summary>
        /// Manages placeholder text for the search box. Clears it on focus.
        /// </summary>
        private void SearchTextBox_Enter(object sender, EventArgs e)
        {
            if (searchTextBox.Text == "Search by Provider...")
            {
                searchTextBox.Text = "";
                searchTextBox.ForeColor = System.Drawing.SystemColors.ControlText;
            }
        }

        /// <summary>
        /// Manages placeholder text for the search box. Restores it if the box is empty on unfocus.
        /// </summary>
        private void SearchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(searchTextBox.Text))
            {
                searchTextBox.Text = "Search by Provider...";
                searchTextBox.ForeColor = System.Drawing.SystemColors.GrayText;
            }
        }

        /// <summary>
        /// Handles the logic for uninstalling one or more selected drivers.
        /// </summary>
        private void UninstallDriverToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one driver to uninstall.", "No Driver Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Build a confirmation string with the names of drivers to be uninstalled.
            string driversToUninstall = "";
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                driversToUninstall += $"- {item.SubItems[0].Text} ({item.SubItems[1].Text})\n"; // e.g., oem123.inf (My Driver)
            }

            // Show a confirmation dialog to the user with a warning about administrator privileges.
            DialogResult confirmResult = MessageBox.Show(
                $"Are you sure you want to uninstall the following driver(s)?\n\n{driversToUninstall}\nThis operation requires administrator privileges and may require a system restart.",
                "Confirm Uninstall",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirmResult == DialogResult.Yes)
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    string publishedName = item.SubItems[0].Text; // Get Published Name (e.g., oemXX.inf)
                    Process p = new Process();
                    p.StartInfo.FileName = Path.Combine(Environment.SystemDirectory, "pnputil.exe");
                    p.StartInfo.Arguments = $"/delete-driver {publishedName} /uninstall";
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;

                    try
                    {
                        p.Start();
                        string output = p.StandardOutput.ReadToEnd();
                        string error = p.StandardError.ReadToEnd();
                        p.WaitForExit();

                        if (p.ExitCode == 0)
                        {
                            MessageBox.Show($"Successfully uninstalled {publishedName}.\nOutput:\n{output}", "Uninstall Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show($"Failed to uninstall {publishedName}.\nError:\n{error}\nOutput:\n{output}", "Uninstall Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred while trying to uninstall {publishedName}:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                // Refresh the driver list after the uninstall operations are complete.
                EnumerateDriversButton_Click(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Filters the cached list of allDrivers based on the search text and updates the ListView.
        /// </summary>
        private void ApplyFilter(string searchText)
        {
            if (allDrivers == null || !allDrivers.Any())
            {
                MessageBox.Show("Please enumerate drivers first by clicking 'Enumerate Drivers' button.", "No Drivers Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                listView1.Items.Clear();
                return;
            }

            // Filter the master list of drivers. The search is case-insensitive and targets the Provider property.
            // An empty or placeholder search text will show all drivers.
            var filteredDrivers = allDrivers.Where(d =>
                searchText.Trim()== "Search by Provider..." ||string.IsNullOrEmpty(searchText.Trim()) || 
                (d.Provider != null && d.Provider.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();

            PopulateListView(filteredDrivers);
            SortListView();
        }

        /// <summary>
        /// Clears and populates the ListView with a given list of drivers.
        /// </summary>
        private void PopulateListView(List<DriverInfo> driversToDisplay)
        {
            listView1.Items.Clear();
            foreach (var driver in driversToDisplay)
            {
                var item = new ListViewItem(driver.PublishedName);
                item.SubItems.Add(driver.OriginalName);
                item.SubItems.Add(driver.Provider);
                item.SubItems.Add(driver.DriverVersion);
                item.SubItems.Add(driver.Class);
                listView1.Items.Add(item);
            }
        }

        /// <summary>
        /// Triggers the ListView to re-sort its items based on the current sorter settings.
        /// </summary>
        private void SortListView()
        {
            this.listView1.Sort();
        }
    }

    /// <summary>
    /// An IComparer implementation for sorting ListView items by a specific column.
    /// </summary>
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
            ListViewItem listviewX = (ListViewItem)x;
            ListViewItem listviewY = (ListViewItem)y;

            // Compare the text of the specified column.
            int compareResult = ObjectCompare.Compare(listviewX.SubItems[ColumnToSort].Text, listviewY.SubItems[ColumnToSort].Text);

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

        /// <summary>
        /// The column to be sorted.
        /// </summary>
        public int SortColumn
        {
            set { ColumnToSort = value; }
            get { return ColumnToSort; }
        }

        /// <summary>
        /// The order of sorting (Ascending, Descending, or None).
        /// </summary>
        public SortOrder Order
        {
            set { OrderOfSort = value; }
            get { return OrderOfSort; }
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
