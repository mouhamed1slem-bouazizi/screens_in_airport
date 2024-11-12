using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Renci.SshNet;
using System.Windows.Forms;
using System.Drawing;

namespace screens_in_airport
{
    public partial class Form1 : Form
    {
        private List<HostInfo> hosts = new List<HostInfo>();
        private string excelFilePath = @"C:\path\to\your\excel\file.xlsx";
        private TextBox pathTextBox;
        private Button browseButton;
        private ListView hostListView;
        private TextBox commandBox;
        private Button executeButton;
        private Label resultLabel;
        private TableLayoutPanel mainTableLayout;
        private Panel topPanel;
        private Panel bottomPanel;
        private Panel leftPanel;
        private Panel rightPanel;

        public class HostInfo
        {
            public string Hostname { get; set; }
            public string IpAddress { get; set; }
            public string Username { get; set; } = "root";
        }

        public Form1()
        {
            InitializeComponents();
            this.WindowState = FormWindowState.Maximized;
        }

        private void InitializeComponents()
        {
            this.Size = new System.Drawing.Size(1024, 768);
            this.Text = "Network Manager";
            this.MinimumSize = new System.Drawing.Size(800, 600);

            // Create main table layout with proper docking
            mainTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            // Set column and row styles for proper scaling
            mainTableLayout.ColumnStyles.Clear();
            mainTableLayout.RowStyles.Clear();
            mainTableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            mainTableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Top panel with proper anchoring
            topPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
                AutoSize = true
            };

            Label pathLabel = new Label
            {
                Text = "Excel File Path:",
                AutoSize = true,
                Location = new Point(5, 8)
            };

            pathTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Height = 23,
                Margin = new Padding(5, 5, 85, 5)
            };

            browseButton = new Button
            {
                Text = "Browse",
                Width = 75,
                Height = 23,
                Dock = DockStyle.Right
            };
            browseButton.Click += BrowseButton_Click;

            // Left panel setup
            leftPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            Label hostsLabel = new Label
            {
                Text = "Connected Hosts:",
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(0, 0, 0, 5)
            };

            hostListView = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false,
                Dock = DockStyle.Fill
            };

            hostListView.Columns.Add("Name", -2);
            hostListView.Columns.Add("IP Address", -2);

            // Right panel setup
            rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            TableLayoutPanel rightTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                AutoSize = true
            };

            rightTableLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 65F));
            rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label commandLabel = new Label
            {
                Text = "Command:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 0, 0, 5)
            };

            commandBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9.75F)
            };

            executeButton = new Button
            {
                Text = "Execute Command",
                Dock = DockStyle.Fill,
                Height = 30
            };
            executeButton.Click += ExecuteButton_Click;

            resultLabel = new Label
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9.75F),
                AutoSize = false
            };

            // Add controls to right table layout
            rightTableLayout.Controls.Add(commandLabel, 0, 0);
            rightTableLayout.Controls.Add(commandBox, 0, 1);
            rightTableLayout.Controls.Add(executeButton, 0, 2);
            rightTableLayout.Controls.Add(resultLabel, 0, 3);

            // Add controls to panels
            topPanel.Controls.Add(browseButton);
            topPanel.Controls.Add(pathTextBox);
            topPanel.Controls.Add(pathLabel);

            leftPanel.Controls.Add(hostListView);
            leftPanel.Controls.Add(hostsLabel);

            rightPanel.Controls.Add(rightTableLayout);

            // Add panels to main table layout
            mainTableLayout.Controls.Add(topPanel, 0, 0);
            mainTableLayout.SetColumnSpan(topPanel, 2);
            mainTableLayout.Controls.Add(leftPanel, 0, 1);
            mainTableLayout.Controls.Add(rightPanel, 1, 1);

            // Add table layout to form
            this.Controls.Add(mainTableLayout);

            // Handle resize events
            this.Resize += MainForm_Resize;
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            // Update path textbox width
            pathTextBox.Width = topPanel.Width - browseButton.Width - 120;
            browseButton.Location = new Point(pathTextBox.Right + 5, pathTextBox.Location.Y);

            // Auto-size ListView columns
            hostListView.Columns[0].Width = (int)(hostListView.Width * 0.45);
            hostListView.Columns[1].Width = (int)(hostListView.Width * 0.45);

            // Force layout update
            mainTableLayout.PerformLayout();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                    pathTextBox.Text = excelFilePath;
                    LoadExcelData();
                }
            }
        }

        private void LoadExcelData()
        {
            try
            {
                hosts.Clear();
                hostListView.Items.Clear();
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var host = new HostInfo
                        {
                            Hostname = worksheet.Cells[row, 1].Value?.ToString()?.Trim(),
                            IpAddress = worksheet.Cells[row, 2].Value?.ToString()?.Trim()
                        };

                        if (!string.IsNullOrEmpty(host.Hostname) && !string.IsNullOrEmpty(host.IpAddress))
                        {
                            hosts.Add(host);
                            ListViewItem item = new ListViewItem(host.Hostname);
                            item.SubItems.Add(host.IpAddress);
                            item.Tag = host;
                            hostListView.Items.Add(item);
                        }
                    }
                }

                if (hosts.Count > 0)
                {
                    MessageBox.Show($"Successfully loaded {hosts.Count} hosts from Excel file.",
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No valid hosts found in the Excel file. Please check the file format.",
                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteCommand(HostInfo host, string command)
        {
            try
            {
                using (var client = new SshClient(host.IpAddress, "root", "123456"))
                {
                    resultLabel.Text = $"Connecting to {host.Hostname} ({host.IpAddress})...";
                    Application.DoEvents();

                    client.Connect();
                    if (client.IsConnected)
                    {
                        resultLabel.Text = $"Executing command on {host.Hostname} ({host.IpAddress})...";
                        Application.DoEvents();

                        var result = client.RunCommand(command);
                        resultLabel.Text = $"Results for {host.Hostname} ({host.IpAddress}):\n\n{result.Result}";
                        client.Disconnect();
                    }
                    else
                    {
                        resultLabel.Text = $"Failed to connect to {host.Hostname} ({host.IpAddress})";
                    }
                }
            }
            catch (Exception ex)
            {
                resultLabel.Text = $"Error executing command on {host.Hostname} ({host.IpAddress}):\n{ex.Message}";
            }
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            if (hostListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select a host first.", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(commandBox.Text))
            {
                MessageBox.Show("Please enter a command.", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedHost = (HostInfo)hostListView.SelectedItems[0].Tag;
            ExecuteCommand(selectedHost, commandBox.Text);
        }
    }
}