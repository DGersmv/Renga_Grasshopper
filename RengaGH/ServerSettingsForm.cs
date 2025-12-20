/*  Server Settings Form for Renga Plugin
 *
 *  Form for configuring TCP server port and managing server state
 *
 *  Copyright Renga Software LLC, 2025. All rights reserved.
 */

#nullable disable
using System;
using System.Drawing;
using System.Windows.Forms;

namespace RengaPlugin
{
    public class ServerSettingsForm : Form
    {
        private TextBox portTextBox;
        private Button startStopButton;
        private Label statusLabel;
        private Label portLabel;
        private bool isServerRunning;

        public int Port { get; private set; } = 50100;
        public bool IsServerRunning => isServerRunning;
        public event EventHandler? StartServerRequested;
        public event EventHandler? StopServerRequested;
        public event EventHandler<int>? PortChanged;

        public ServerSettingsForm(int currentPort, bool serverRunning)
        {
            Port = currentPort;
            isServerRunning = serverRunning;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Standard Renga colors
            // Background: SystemColors.Control (light gray)
            // Text: SystemColors.ControlText (dark text)
            // Buttons: SystemColors.ButtonFace with standard button styling

            this.SuspendLayout();

            // Form properties
            this.Text = "Renga_Sverchok Server Settings";
            this.Size = new Size(350, 180);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = SystemColors.Control;
            this.Font = new Font("Segoe UI", 9F);

            // Port Label
            portLabel = new Label
            {
                Text = "TCP Port:",
                Location = new Point(20, 25),
                Size = new Size(80, 23),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.ControlText
            };
            this.Controls.Add(portLabel);

            // Port TextBox
            portTextBox = new TextBox
            {
                Text = Port.ToString(),
                Location = new Point(110, 23),
                Size = new Size(100, 23),
                TextAlign = HorizontalAlignment.Right
            };
            portTextBox.TextChanged += PortTextBox_TextChanged;
            portTextBox.KeyPress += PortTextBox_KeyPress;
            this.Controls.Add(portTextBox);

            // Status Label
            statusLabel = new Label
            {
                Text = isServerRunning ? "Status: Running" : "Status: Stopped",
                Location = new Point(20, 60),
                Size = new Size(300, 23),
                AutoSize = false,
                ForeColor = isServerRunning ? Color.Green : Color.Gray,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            this.Controls.Add(statusLabel);

            // Start/Stop Button
            startStopButton = new Button
            {
                Text = isServerRunning ? "Stop Server" : "Start Server",
                Location = new Point(20, 100),
                Size = new Size(120, 30),
                UseVisualStyleBackColor = true,
                BackColor = SystemColors.ButtonFace,
                ForeColor = SystemColors.ControlText,
                FlatStyle = FlatStyle.Standard
            };
            startStopButton.Click += StartStopButton_Click;
            this.Controls.Add(startStopButton);

            // Close Button
            var closeButton = new Button
            {
                Text = "Close",
                Location = new Point(250, 100),
                Size = new Size(70, 30),
                UseVisualStyleBackColor = true,
                BackColor = SystemColors.ButtonFace,
                ForeColor = SystemColors.ControlText,
                DialogResult = DialogResult.OK
            };
            this.Controls.Add(closeButton);

            this.ResumeLayout(false);
        }

        private void PortTextBox_KeyPress(object? sender, KeyPressEventArgs e)
        {
            // Allow only digits and backspace
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void PortTextBox_TextChanged(object? sender, EventArgs e)
        {
            if (int.TryParse(portTextBox.Text, out int port))
            {
                if (port >= 1024 && port <= 65535)
                {
                    Port = port;
                    PortChanged?.Invoke(this, port);
                }
            }
        }

        private void StartStopButton_Click(object? sender, EventArgs e)
        {
            if (isServerRunning)
            {
                StopServerRequested?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                // Validate port before starting
                if (!int.TryParse(portTextBox.Text, out int port) || port < 1024 || port > 65535)
                {
                    MessageBox.Show(
                        "Invalid port number. Port must be between 1024 and 65535.",
                        "Invalid Port",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                StartServerRequested?.Invoke(this, EventArgs.Empty);
            }
        }

        public void UpdateServerStatus(bool running)
        {
            isServerRunning = running;
            if (statusLabel != null)
            {
                statusLabel.Text = running ? "Status: Running" : "Status: Stopped";
                statusLabel.ForeColor = running ? Color.Green : Color.Gray;
            }
            if (startStopButton != null)
            {
                startStopButton.Text = running ? "Stop Server" : "Start Server";
            }
        }
    }
}

