// ========================== MainForm.cs ==========================
using InvoiceGenerator;               // flat namespace (models + services in root)
using Microsoft.VisualBasic;          // for Interaction.InputBox
using Microsoft.VisualBasic.FileIO;   // for TextFieldParser
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using InvoiceGenerator.Models;
using InvoiceGenerator.Services;



namespace InvoiceGenerator
{
    public class MainForm : Form
    {
        // === FOLDER STRUCTURE SETTINGS ===
        private const string PreferredRoot = @"C:\Mehringer Construction Invoice";
        private const int SingleLineH = 50;
        private const decimal MaterialMarkupRate = 0.10m;

        // === PATHS ===
        private readonly string RootFolder;
        private readonly string CompanyFolder;
        private readonly string GeneratedFolder;
        private readonly string KrempFolder;
        private readonly string EmployeeHoursFolder;
        private readonly string CompanyInfoPath;
        private readonly string LogoPath;
        private readonly string HomeDepotFolder;

        // === DATA ===
        private CompanyInfo _company = new CompanyInfo();
        private readonly BindingList<InvoiceItem> _items = new();
        private readonly HashSet<string> _importedFiles = new(StringComparer.OrdinalIgnoreCase);

        // === UI ===
        private PictureBox pbLogo = null!;
        private TextBox txtCustFirst = null!, txtCustLast = null!, txtBillTo = null!,
                        txtProjectNameOrNumber = null!, txtProjectTag = null!, txtInvoiceNumber = null!;
        private DateTimePicker dtpDate = null!;
        private Label lblDueDate = null!;
        private NumericUpDown nudTaxRate = null!, nudDefaultLaborRate = null!;
        private DataGridView grid = null!;

        private ToolStrip tool = null!;
        private ToolStripButton tbtnLoadPdf = null!, tbtnLoadHours = null!, tbtnAddSingleLabor = null!,
                                tbtnAddRow = null!, tbtnRemoveRow = null!, tbtnGenerate = null!,
                                tbtnLoadHomeDepotCsv = null!;
        private ToolStripControlHost chkFilterHost = null!;
        private CheckBox chkFilterProject = null!;

        private Label lblSubtotal = null!, lblMarkup = null!, lblTax = null!, lblTotal = null!;

        public MainForm()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 10);
            MinimumSize = new Size(1200, 800);
            WindowState = FormWindowState.Maximized;
            Text = "Mehringer Construction - Invoice Generator";

            // Resolve root folder
            if (Directory.Exists(PreferredRoot))
                RootFolder = PreferredRoot;
            else
            {
                var exe = AppContext.BaseDirectory;
                var proj = Directory.GetParent(exe)?.Parent?.Parent?.Parent?.FullName
                           ?? Directory.GetParent(exe)!.FullName;
                RootFolder = Path.GetFullPath(Path.Combine(proj, @"..\")); // repo root
            }

            // Paths
            CompanyFolder = Path.Combine(RootFolder, "Company Information");
            GeneratedFolder = Path.Combine(RootFolder, "Generated Invoices");
            KrempFolder = Path.Combine(RootFolder, "Kremp Invoices");
            EmployeeHoursFolder = Path.Combine(RootFolder, "Employee Hours");
            CompanyInfoPath = Path.Combine(CompanyFolder, "Company information.txt");
            LogoPath = Path.Combine(CompanyFolder, "Mehringer logo.png");
            HomeDepotFolder = Path.Combine(RootFolder, "Home Depot Invoices");

            Directory.CreateDirectory(GeneratedFolder);
            Directory.CreateDirectory(EmployeeHoursFolder);
            Directory.CreateDirectory(HomeDepotFolder);

            BuildUi();
            LoadCompany();
            LoadLogo();

            AutoSuggestInvoiceNumber();
            UpdateDueDate();
            AutoSuggestProjectTag();

            // Reactive helpers
            dtpDate.ValueChanged += (_, __) => { UpdateDueDate(); AutoSuggestProjectTag(); AutoSuggestInvoiceNumber(); };
            txtCustLast.TextChanged += (_, __) => AutoSuggestProjectTag();
            txtProjectNameOrNumber.TextChanged += (_, __) => AutoSuggestProjectTag();

            UpdateTotals();
        }

        // ---------------- UI ----------------
        private void BuildUi()
        {
            tool = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                ImageScalingSize = new Size(20, 20),
                RenderMode = ToolStripRenderMode.System,
                LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow,
                AutoSize = false,
                Height = 120,
                Padding = new Padding(12, 14, 12, 14),
                Dock = DockStyle.Top
            };

            ToolStripButton MakeBtn(string text, EventHandler onClick, int width)
            {
                var b = new ToolStripButton(text)
                {
                    AutoSize = false,
                    Width = width,
                    Height = 64,
                    Padding = new Padding(12, 8, 12, 8),
                    Margin = new Padding(12, 6, 14, 6),
                    DisplayStyle = ToolStripItemDisplayStyle.Text,
                    TextAlign = ContentAlignment.MiddleCenter
                };
                b.Click += onClick;
                return b;
            }
            ToolStripItem Gap(int w = 24) => new ToolStripLabel { AutoSize = false, Width = w };
            ToolStripItem Line(int sideGap = 12) => new ToolStripSeparator { AutoSize = false, Width = 1, Margin = new Padding(sideGap, 12, sideGap, 12) };

            chkFilterProject = new CheckBox
            {
                Text = "Select Current Project Hours",
                Checked = true,
                AutoSize = false,
                Appearance = Appearance.Button,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Standard,
                Padding = new Padding(12, 6, 12, 6),
                MinimumSize = new Size(280, 100),
                Size = new Size(280, 100),
                Margin = new Padding(0)
            };
            chkFilterHost = new ToolStripControlHost(chkFilterProject)
            {
                AutoSize = false,
                Margin = new Padding(12, 4, 20, 4),
                Size = new Size(300, 48)
            };

            tbtnLoadPdf = MakeBtn("Upload Supplier PDF", BtnLoadPdf_Click, 240);
            tbtnLoadHomeDepotCsv = MakeBtn("Upload Home Depot CSV", BtnLoadHomeDepotCsv_Click, 280);
            tbtnLoadHours = MakeBtn("Upload Employee Hours", BtnLoadHours_Click, 300);
            tbtnAddSingleLabor = MakeBtn("Single Labor Entry", BtnAddSingleLabor_Click, 220);
            tbtnAddRow = MakeBtn("Manual Item", (s, e) => { _items.Add(new InvoiceItem()); UpdateTotals(); }, 180);
            tbtnRemoveRow = MakeBtn("Remove Selected", (s, e) => { if (grid?.CurrentRow?.DataBoundItem is InvoiceItem it) _items.Remove(it); UpdateTotals(); }, 200);
            tbtnGenerate = MakeBtn("Create Customer Invoice", BtnGenerate_Click, 320);
            tbtnGenerate.Alignment = ToolStripItemAlignment.Right;

            tool.Items.Add(chkFilterHost);
            tool.Items.Add(Gap());
            tool.Items.Add(tbtnLoadPdf);
            tool.Items.Add(Gap());
            tool.Items.Add(tbtnLoadHomeDepotCsv);
            tool.Items.Add(Gap());
            tool.Items.Add(tbtnLoadHours);
            tool.Items.Add(Gap());
            tool.Items.Add(tbtnAddSingleLabor);
            tool.Items.Add(Line());
            tool.Items.Add(tbtnAddRow);
            tool.Items.Add(Gap());
            tool.Items.Add(tbtnRemoveRow);
            tool.Items.Add(tbtnGenerate);
            tool.Items.Insert(tool.Items.IndexOf(tbtnGenerate), Gap(24));

            // ===== Header (logo + inputs) =====
            var headerPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12),
                ColumnCount = 2
            };
            headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            pbLogo = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle
            };
            AddAt(headerPanel, pbLogo, 0, 0);

            var inputs = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
                AutoSize = true
            };
            inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));
            inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            AddAt(headerPanel, inputs, 1, 0);

            // Row 0
            txtInvoiceNumber = LabeledText(inputs, "Invoice #", 0, 0);
            dtpDate = new DateTimePicker { Dock = DockStyle.Fill };
            AddAt(inputs, Labeled("Invoice Date", dtpDate, SingleLineH, false), 1, 0);

            lblDueDate = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            AddAt(inputs, Labeled("Due Date (+30d)", lblDueDate, SingleLineH, false), 2, 0);

            nudTaxRate = new NumericUpDown { DecimalPlaces = 2, Value = 7, Width = 100 };
            AddAt(inputs, Labeled("Tax Rate (%)", nudTaxRate, SingleLineH, false), 3, 0);

            // Row 1
            txtCustFirst = LabeledText(inputs, "Customer First", 0, 1);
            txtCustLast = LabeledText(inputs, "Customer Last", 1, 1);
            txtProjectNameOrNumber = LabeledText(inputs, "Project Name/Number", 3, 1);

            // Row 2 — full-width
            txtBillTo = LabeledText(inputs, "Billing Address", 0, 2, multiLine: true, multiHeight: 200);
            inputs.SetColumnSpan(txtBillTo.Parent!, 4);

            // Row 3
            txtProjectTag = LabeledText(inputs, "Project Tag", 0, 3);
            nudDefaultLaborRate = new NumericUpDown
            {
                DecimalPlaces = 2,
                Width = 400,
                Minimum = 0,
                Maximum = 1000,
                Increment = 0.25m,
                Value = 41.50m
            };
            AddAt(inputs, Labeled("Default Labor Rate", nudDefaultLaborRate, 28, false), 1, 3);

            // ===== Totals panel =====
            var totals = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                ColumnCount = 2,
                Padding = new Padding(12, 0, 12, 16),
                Margin = new Padding(0)
            };
            totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            totals.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300));

            var totalsInner = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = false,
                Padding = new Padding(8, 2, 8, 6)
            };
            totalsInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            AddAt(totals, new Panel(), 0, 0);
            AddAt(totals, totalsInner, 1, 0);

            lblSubtotal = TotalsRow(totalsInner, "Subtotal:");
            lblMarkup = TotalsRow(totalsInner, "Material Markup (10%):");
            lblTax = TotalsRow(totalsInner, "Tax:");
            lblTotal = TotalsRow(totalsInner, "Total:", bold: true, bigger: true);

            // ===== Grid =====
            grid = new DataGridView
            {
                AutoGenerateColumns = false,
                DataSource = _items,
                AllowUserToAddRows = false,
                BackgroundColor = Color.White,
                Dock = DockStyle.Fill,
                ColumnHeadersVisible = true,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 64
            };
            grid.RowTemplate.Height = 64;
            grid.DefaultCellStyle.Padding = new Padding(0, 4, 0, 4);

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Product/Service", DataPropertyName = "ProductOrService", Width = 220 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Description", DataPropertyName = "Description", Width = 600 });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Qty",
                DataPropertyName = "Quantity",
                Width = 150,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "N2" }
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Rate",
                DataPropertyName = "Rate",
                Width = 150,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "C2" }
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Amount",
                DataPropertyName = "LineTotal",
                Width = 150,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Format = "C2" }
            });

            grid.CellEndEdit += (_, __) => UpdateTotals();
            grid.RowsAdded += (_, __) => UpdateTotals();
            grid.RowsRemoved += (_, __) => UpdateTotals();

            var gridBox = new GroupBox { Text = "Line Items", Dock = DockStyle.Fill, Padding = new Padding(8) };
            var gridScrollHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(0) };
            gridScrollHost.Controls.Add(grid);
            gridBox.Controls.Add(gridScrollHost);

            // Split center into Grid (top) + Totals (bottom)
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.Panel2,
                IsSplitterFixed = false,
                SplitterWidth = 6
            };
            split.Panel2MinSize = 150;
            split.Panel1.Controls.Add(gridBox);
            split.Panel2.Controls.Add(totals);

            // Main body
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0)
            };
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            body.Controls.Add(headerPanel, 0, 0);
            body.Controls.Add(split, 0, 1);

            // Add to form with a spacer
            SuspendLayout();
            var spacer = new Panel { Dock = DockStyle.Top, Height = 16 };
            Controls.Add(body);
            Controls.Add(spacer);
            Controls.Add(tool);
            ResumeLayout(true);
        }

        // ---------- Helpers ----------
        private static string MakeSafeFilePart(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "Unknown";
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(s.Trim().Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            cleaned = Regex.Replace(cleaned, @"[\s_]+", "_");
            return cleaned.Length == 0 ? "Unknown" : cleaned;
        }
        private static bool IsReturnOrCreditPdf(string pdfPath)
        {
            try
            {
                using var pdf = PdfDocument.Open(pdfPath);
                foreach (var page in pdf.GetPages().Take(2))
                {
                    var text = ContentOrderTextExtractor.GetText(page) ?? string.Empty;
                    var up = text.ToUpperInvariant();
                    if (up.Contains("RETURN") ||
                        up.Contains("CREDIT MEMO") ||
                        up.Contains("MERCHANDISE RETURN") ||
                        up.Contains("CREDIT INVOICE") ||
                        up.Contains("CREDIT #") ||
                        up.Contains("CUSTOMER CREDIT") ||
                        up.Contains("REFUND"))
                        return true;
                }
            }
            catch { /* fall through */ }
            return false;
        }

        private static void AddAt(TableLayoutPanel parent, Control control, int col, int row)
        {
            parent.Controls.Add(control);
            parent.SetColumn(control, col);
            parent.SetRow(control, row);
        }

        private static Control Labeled(string caption, Control inner, int controlHeight = 100, bool multiline = false)
        {
            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                Padding = new Padding(2)
            };
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, controlHeight));

            var lbl = new Label { Text = caption, Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 0, 0, 4) };

            if (multiline && inner is TextBox tb)
            {
                tb.Multiline = true;
                tb.ScrollBars = ScrollBars.Vertical;
                inner.MinimumSize = new Size(0, controlHeight);
            }

            inner.Dock = DockStyle.Fill;
            tlp.Controls.Add(lbl, 0, 0);
            tlp.Controls.Add(inner, 0, 1);
            return tlp;
        }

        private TextBox LabeledText(Control parent, string caption, int col, int row,
                                    bool multiLine = false, int multiHeight = 90)
        {
            var tb = new TextBox { Dock = DockStyle.Fill, Multiline = multiLine };
            var container = Labeled(caption, tb, multiLine ? multiHeight : SingleLineH, multiLine);
            AddAt((TableLayoutPanel)parent, container, col, row);
            return tb;
        }

        private static Label TotalsRow(Control parent, string label, bool bold = false, bool bigger = false)
        {
            var row = new Panel { Dock = DockStyle.Top, Height = bigger ? 44 : 32 };

            var left = new Label
            {
                Text = label,
                Dock = DockStyle.Left,
                Width = 140,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var right = new Label
            {
                Text = "$0.00",
                Dock = DockStyle.Right,
                Width = 140,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("Segoe UI", bigger ? 11 : 9, bold ? FontStyle.Bold : FontStyle.Regular)
            };

            row.Controls.Add(right);
            row.Controls.Add(left);
            parent.Controls.Add(row);
            return right;
        }

        // ---------- Data load ----------
        private void LoadCompany()
        {
            try
            {
                _company = CompanyInfoLoader.LoadFromFile(CompanyInfoPath);
                Text = $"{_company.Name} – Invoice Generator";
            }
            catch
            {
                _company = new CompanyInfo { Name = "Mehringer's Construction LLC" };
            }
        }

        // Try to find a usable logo file
        private string? FindLogoPath()
        {
            if (File.Exists(LogoPath)) return LogoPath;
            var patterns = new[]
            {
                "Mehringer logo*.png", "Mehringer*logo*.png", "logo*.png",
                "Mehringer logo*.jpg", "Mehringer*logo*.jpg", "logo*.jpg",
                "Mehringer logo*.jpeg","Mehringer*logo*.jpeg","logo*.jpeg"
            };
            foreach (var p in patterns)
            {
                var hit = Directory
                    .EnumerateFiles(CompanyFolder, p, System.IO.SearchOption.TopDirectoryOnly)
                    .OrderBy(f => f)
                    .FirstOrDefault();
                if (hit != null) return hit;
            }
            return null;
        }

        private void LoadLogo()
        {
            try
            {
                var path = FindLogoPath();
                if (path == null) return;

                var bytes = File.ReadAllBytes(path);
                using (var ms = new MemoryStream(bytes))
                using (var img = Image.FromStream(ms))
                {
                    pbLogo.Image = new Bitmap(img);
                }
            }
            catch { /* optional */ }
        }

        // ---------- Auto helpers ----------
        private void UpdateDueDate() =>
            lblDueDate.Text = dtpDate.Value.Date.AddDays(30).ToShortDateString();

        private void AutoSuggestProjectTag()
        {
            var yy = dtpDate.Value.ToString("yy");
            var last = txtCustLast.Text.Trim();
            var proj = txtProjectNameOrNumber.Text.Trim();
            txtProjectTag.Text = string.IsNullOrWhiteSpace(last)
                ? yy
                : (string.IsNullOrWhiteSpace(proj) ? $"{yy}-{last}" : $"{yy}-{last}/{proj}");
        }

        private void AutoSuggestInvoiceNumber()
        {
            var yy = dtpDate.Value.ToString("yy", CultureInfo.InvariantCulture);
            int existingCount;
            try
            {
                existingCount = Directory
                    .EnumerateFiles(GeneratedFolder, "*.pdf", System.IO.SearchOption.TopDirectoryOnly
)
                    .Count();
            }
            catch { existingCount = 0; }
            txtInvoiceNumber.Text = $"{yy}-{(existingCount + 1):000}";
        }

        private static (string first, string last) SplitFirstLast(string name)
        {
            var parts = (name ?? string.Empty).Trim()
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return (string.Empty, string.Empty);
            if (parts.Length == 1) return (parts[0], string.Empty);
            return (string.Join(" ", parts.Take(parts.Length - 1)), parts.Last());
        }

        // Push imported header values to UI
        private void ApplyImportedHeaderToUi(ImportedInvoice imported)
        {
            if (imported == null) return;

            // Invoice number/date/due
            if (!string.IsNullOrWhiteSpace(imported.Header.InvoiceNumber))
                txtInvoiceNumber.Text = imported.Header.InvoiceNumber;

            if (imported.Header.InvoiceDate.HasValue)
                dtpDate.Value = imported.Header.InvoiceDate.Value;

            var due = imported.Header.DueDate ?? dtpDate.Value.AddDays(30);
            lblDueDate.Text = due.ToShortDateString();

            // Bill To (only if it doesn't look like a label block)
            if (!string.IsNullOrWhiteSpace(imported.Header.BillTo))
            {
                var billToUpper = imported.Header.BillTo.ToUpperInvariant();
                var looksLikeLabelBlock =
                    billToUpper.Contains("JOB NO:") ||
                    billToUpper.Contains("PURCHASE ORDER:") ||
                    billToUpper.Contains("REFERENCE:") ||
                    billToUpper.Contains("TERMS:") ||
                    billToUpper.Contains("CLERK:") ||
                    billToUpper.Contains("DATE / TIME:");

                if (!looksLikeLabelBlock)
                    txtBillTo.Text = imported.Header.BillTo;
            }

            // Customer name → only if it looks like a name (no colon, not all-caps)
            if (!string.IsNullOrWhiteSpace(imported.Header.CustomerName))
            {
                string name = imported.Header.CustomerName.Trim();
                bool plausibleName =
                    !name.Contains(":") &&
                    !(name.ToUpperInvariant() == name && name.Any(char.IsLetter));

                if (plausibleName)
                {
                    var parts = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 1)
                    {
                        txtCustFirst.Text = parts[0];
                    }
                    else if (parts.Length >= 2)
                    {
                        txtCustFirst.Text = string.Join(" ", parts.Take(parts.Length - 1));
                        txtCustLast.Text = parts.Last();
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(imported.Header.ProjectTag))
                txtProjectTag.Text = imported.Header.ProjectTag;

            if (!string.IsNullOrWhiteSpace(imported.Header.ProjectName))
                txtProjectNameOrNumber.Text = imported.Header.ProjectName;

            if (imported.Header.TaxRatePercent > 0)
                nudTaxRate.Value = Math.Min(nudTaxRate.Maximum, Math.Max(nudTaxRate.Minimum, imported.Header.TaxRatePercent));

            AutoSuggestInvoiceNumber();
            UpdateDueDate();
            AutoSuggestProjectTag();
        }


        // ---------- Buttons ----------
        // PDF → CSV → Parse → Append (append-only, de-duped, supports returns)
        private void BtnLoadPdf_Click(object? s, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                InitialDirectory = Directory.Exists(KrempFolder)
                    ? KrempFolder
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Title = "Select KREMPP Supplier Invoice (PDF)"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            var pdfPath = ofd.FileName;

            // Session duplicate guard (optional)
            if (_importedFiles.Contains(pdfPath))
            {
                var ans = MessageBox.Show(
                    "You already imported this PDF in this session.\n\nImport again anyway?",
                    "Duplicate File",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (ans != DialogResult.Yes) return;
            }

            // 1) Convert PDF → CSV (saved automatically into "CSV Invoice Files")
            string csvPath;

            try
            {
                // Convert PDF → CSV
                PdfToCsvConverter.Convert(pdfPath);

                // Build the CSV path that the converter just wrote
                csvPath = Path.Combine(
                    Path.GetDirectoryName(pdfPath) ?? ".",
                    "CSV Invoice Files",
                    Path.GetFileNameWithoutExtension(pdfPath) + ".csv"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to convert PDF to CSV:\n{ex.Message}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }



            // 2) Parse CSV → ImportedInvoice
            ImportedInvoice inv;
            try
            {
                if (!File.Exists(csvPath))
                {
                    MessageBox.Show($"CSV not found:\n{csvPath}", "Import Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                inv = CsvInvoiceParser.Parse(csvPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to read CSV:\n{ex.Message}",
                    "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (inv?.Items == null || inv.Items.Count == 0)
            {
                MessageBox.Show(
                    $"No line items were found in:\n{Path.GetFileName(pdfPath)}",
                    "PDF→CSV Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3) Return/credit detection (forces negative qty/amount when appropriate)
            bool looksLikeReturn =
                Path.GetFileName(pdfPath).IndexOf("return", StringComparison.OrdinalIgnoreCase) >= 0
                || inv.Items.Any(x =>
                       (x.Description ?? "").IndexOf("CREDIT", StringComparison.OrdinalIgnoreCase) >= 0
                    || (x.Description ?? "").IndexOf("RETURN", StringComparison.OrdinalIgnoreCase) >= 0);

            bool hasNegativeAlready = inv.Items.Any(x => x.Qty < 0 || x.Amount < 0);
            bool isReturn = looksLikeReturn || hasNegativeAlready;

            // 4) De-dupe against current grid
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var it in _items)
            {
                existing.Add($"{it.Description}|{it.Quantity.ToString(CultureInfo.InvariantCulture)}|{it.Rate.ToString(CultureInfo.InvariantCulture)}");
            }

            // De-dupe within this one import
            var perImport = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int added = 0, skippedDup = 0;

            foreach (var r in inv.Items)
            {
                // Normalize values
                decimal qty = r.Qty;
                decimal rate = Math.Abs(r.Rate);  // keep unit price positive
                decimal amount = r.Amount;

                // If it's a return, ensure negatives for qty/amount when non-zero
                if (isReturn)
                {
                    if (amount > 0) amount = -amount;
                    if (qty > 0) qty = -qty;
                }

                // Recover qty from amount/rate if needed
                if (qty == 0 && amount != 0 && rate > 0)
                {
                    qty = Math.Round(amount / rate, 2, MidpointRounding.AwayFromZero);
                }

                var desc = (r.Description ?? "").Trim();

                // De-dupe signature (no ProductOrService on ImportedItem; use constant "Product")
                string sig = $"{desc}|{qty.ToString(CultureInfo.InvariantCulture)}|{rate.ToString(CultureInfo.InvariantCulture)}";

                if (!perImport.Add(sig) || existing.Contains(sig))
                {
                    skippedDup++;
                    continue;
                }

                _items.Add(new InvoiceItem
                {
                    ProductOrService = "Product", // constant since ImportedItem has no ProductOrService
                    Description = desc,
                    Quantity = qty,        // can be negative for returns
                    Rate = rate
                });

                added++;
            }

            _importedFiles.Add(pdfPath);
            UpdateTotals();

            MessageBox.Show(
                $"Imported {added} line item(s) from:\n{Path.GetFileName(pdfPath)}" +
                (skippedDup > 0 ? $"\nSkipped {skippedDup} duplicate row(s)." : ""),
                "PDF→CSV Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }





        // Home Depot CSV import (unchanged in spirit, tidied a bit)
        private void BtnLoadHomeDepotCsv_Click(object? s, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                InitialDirectory = HomeDepotFolder,
                Title = "Select Home Depot CSV"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            // ---------- helpers ----------
            static string Norm(string? x)
            {
                if (string.IsNullOrWhiteSpace(x)) return "";
                var t = x.Trim().Replace("\uFEFF", "").ToLowerInvariant();
                var sb = new StringBuilder(t.Length);
                foreach (var ch in t)
                {
                    if (char.IsLetterOrDigit(ch)) sb.Append(ch);
                    else if (char.IsWhiteSpace(ch)) sb.Append(' ');
                }
                return Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
            }

            static bool TryParseDecimal(string? raw, out decimal value)
            {
                value = 0m;
                if (string.IsNullOrWhiteSpace(raw)) return false;
                var t = raw.Trim();
                var neg = t.StartsWith("(") && t.EndsWith(")");
                t = t.Replace("(", "").Replace(")", "").Replace("$", "").Replace(",", "");
                if (decimal.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ||
                    decimal.TryParse(t, NumberStyles.Any, CultureInfo.CurrentCulture, out v))
                {
                    value = neg ? -v : v;
                    return true;
                }
                return false;
            }

            static bool TryParseDate(string? raw, out DateTime d)
            {
                d = default;
                if (string.IsNullOrWhiteSpace(raw)) return false;
                var s1 = raw.Trim();
                return DateTime.TryParse(s1, CultureInfo.InvariantCulture, DateTimeStyles.None, out d) ||
                       DateTime.TryParse(s1, CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
            }

            static int FindCol(string[] header, params string[] aliases)
            {
                var normHdr = header.Select(Norm).ToArray();
                foreach (var alias in aliases)
                {
                    var normAlias = Norm(alias);
                    for (int i = 0; i < normHdr.Length; i++)
                    {
                        if (normHdr[i] == normAlias) return i;
                        if (normHdr[i].Contains(normAlias)) return i;
                    }
                }
                return -1;
            }
            // -----------------------------

            var invoiceTag = txtProjectTag.Text.Trim();
            if (string.IsNullOrEmpty(invoiceTag))
            {
                MessageBox.Show("Please set a Project Tag on the invoice before importing.", "CSV Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var invDate = dtpDate.Value.Date;
            var minDate = invDate.AddDays(-14);
            var maxDate = invDate.AddDays(14);

            var rows = new List<string[]>();
            try
            {
                using var fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                using var parser = new TextFieldParser(fs)
                {
                    TextFieldType = FieldType.Delimited,
                    HasFieldsEnclosedInQuotes = true,
                    TrimWhiteSpace = false
                };
                parser.SetDelimiters(",");

                while (!parser.EndOfData)
                {
                    string[]? f;
                    try { f = parser.ReadFields(); } catch { continue; }
                    if (f != null) rows.Add(f);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading CSV:\n{ex.Message}", "CSV Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (rows.Count == 0)
            {
                MessageBox.Show("CSV appears to be empty.", "CSV Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Try to locate the real header row within the first 25 rows.
            int headerRowIndex = -1;
            int idxDate = -1, idxJob = -1, idxQty = -1, idxRate = -1, idxDesc = -1;
            int scanLimit = Math.Min(25, rows.Count);

            for (int r = 0; r < scanLimit; r++)
            {
                var header = rows[r];
                var d = FindCol(header, "date", "transaction date", "order date", "purchase date", "trans date");
                var j = FindCol(header, "job name", "project tag", "job", "jobname");
                var q = FindCol(header, "quantity", "qty");
                var pr = FindCol(header, "net unit price", "unit price", "price", "unitprice", "netunitprice");
                if (d >= 0 && j >= 0 && q >= 0 && pr >= 0)
                {
                    headerRowIndex = r;
                    idxDate = d; idxJob = j; idxQty = q; idxRate = pr;
                    idxDesc = FindCol(header, "sku description", "description", "item description", "product description");
                    break;
                }
            }

            // Fallback to known Home Depot layout
            bool usingFallback = false;
            if (headerRowIndex < 0)
            {
                if (rows.Count > 7)
                {
                    headerRowIndex = 6;
                    var header = rows[headerRowIndex];
                    if (header.Length >= 17)
                    {
                        idxDate = 0; idxJob = 4; idxDesc = 6; idxQty = 7; idxRate = 16;
                        usingFallback = true;
                    }
                }
            }

            if (headerRowIndex < 0 || idxDate < 0 || idxJob < 0 || idxQty < 0 || idxRate < 0)
            {
                MessageBox.Show("Could not locate required columns in CSV header.", "CSV Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int rowsRead = 0, added = 0, filteredJob = 0, filteredDate = 0, badNumeric = 0;
            for (int r = headerRowIndex + 1; r < rows.Count; r++)
            {
                var f = rows[r];
                if (f == null || f.Length <= Math.Max(idxRate, Math.Max(idxQty, Math.Max(idxJob, idxDate)))) continue;
                rowsRead++;

                var job = (f[idxJob] ?? "").Trim();
                if (!job.Equals(invoiceTag, StringComparison.OrdinalIgnoreCase))
                {
                    filteredJob++;
                    continue;
                }

                if (!TryParseDate(f[idxDate], out var d) || d.Date < minDate || d.Date > maxDate)
                {
                    filteredDate++;
                    continue;
                }

                if (!TryParseDecimal(f[idxQty], out var qty) ||
                    !TryParseDecimal(f[idxRate], out var rate))
                {
                    badNumeric++;
                    continue;
                }

                string desc = (idxDesc >= 0 && idxDesc < f.Length && !string.IsNullOrWhiteSpace(f[idxDesc]))
                    ? f[idxDesc].Trim()
                    : $"Home Depot Purchase {d:MM/dd/yyyy}";

                _items.Add(new InvoiceItem
                {
                    ProductOrService = "Product",
                    Description = desc,
                    Quantity = qty,
                    Rate = rate
                });
                added++;
            }

            UpdateTotals();

            if (added == 0)
            {
                var mode = usingFallback ? "fallback column positions" : "detected header";
                MessageBox.Show(
                    "No matching rows found for this Project Tag and ±14 day window.\n\n" +
                    $"Mode: {mode}\nHeader row index: {headerRowIndex}\n\n" +
                    $"Rows read: {rowsRead}\n" +
                    $"Filtered by Job: {filteredJob}\n" +
                    $"Filtered by Date: {filteredDate}\n" +
                    $"Bad/Invalid numeric fields: {badNumeric}",
                    "CSV Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    $"Imported {added} line item(s) for Project Tag \"{invoiceTag}\" from {Path.GetFileName(ofd.FileName)}.",
                    "CSV Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnLoadHours_Click(object? s, EventArgs e) => ImportHoursFromPicker(true, true);

        private void BtnAddSingleLabor_Click(object? s, EventArgs e)
        {
            var desc = Interaction.InputBox("Labor description:", "Add Labor", "Labor - All Employees");
            if (string.IsNullOrWhiteSpace(desc)) return;
            if (!decimal.TryParse(Interaction.InputBox("Hours:", "Add Labor", "8"), out var hrs)) return;
            var rate = nudDefaultLaborRate.Value;
            _items.Add(new InvoiceItem { ProductOrService = "Service", Description = desc, Quantity = hrs, Rate = rate });
            UpdateTotals();
        }

        private void BtnGenerate_Click(object? s, EventArgs e)
        {
            if (!_items.Any())
            {
                MessageBox.Show("No items to invoice.");
                return;
            }

            var matSubtotal = _items
    .Where(i => i.ProductOrService.Equals("Product", StringComparison.OrdinalIgnoreCase))
    .Sum(i => i.LineTotal);

            var servSubtotal = _items
                .Where(i => !i.ProductOrService.Equals("Product", StringComparison.OrdinalIgnoreCase))
                .Sum(i => i.LineTotal);

            var markup = Math.Round(matSubtotal * MaterialMarkupRate, 2);

            // For display subtotal
            var subtotalBeforeTax = matSubtotal + servSubtotal + markup;

            // === Tax ONLY on materials + markup ===
            var taxBase = matSubtotal + markup;
            var tax = Math.Round(taxBase * (nudTaxRate.Value / 100m), 2);

            var total = subtotalBeforeTax + tax;


            var lastName = MakeSafeFilePart(txtCustLast.Text);
            var projName = MakeSafeFilePart(txtProjectNameOrNumber.Text);
            var dateStamp = dtpDate.Value.ToString("yyyy-MM-dd");
            var fileName = $"{lastName}_{projName}_{dateStamp}.pdf";
            var outFile = Path.Combine(GeneratedFolder, fileName);

            var invNo = txtInvoiceNumber.Text;
            string? logoForPdf = FindLogoPath() ?? string.Empty;

            PdfExporter.CreateInvoicePdf(
                outFile,
                _company,
                logoForPdf,
                $"{txtCustFirst.Text} {txtCustLast.Text}",
                txtBillTo.Text,
                invNo,
                dtpDate.Value,
                dtpDate.Value.AddDays(30),
                nudTaxRate.Value,
                txtProjectTag.Text,
                txtProjectNameOrNumber.Text,
                _items.ToList(),
                matSubtotal + servSubtotal,
                markup,
                tax,
                total
            );

            MessageBox.Show($"Invoice created:\n{outFile}", "Done");
        }

        private void ImportHoursFromPicker(bool filter, bool annotate)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel/CSV/TXT|*.xlsx;*.csv;*.txt",
                InitialDirectory = EmployeeHoursFolder
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            decimal totalHours;
            if (filter)
            {
                // Strict audit: Column B must equal the invoice tag; sum Column D (HoursImporter normalizes Excel time/day fractions)
                var (sum, matched, skipped, matches) = HoursImporter.AuditSumForInvoice(ofd.FileName, txtProjectTag.Text);
                totalHours = sum;

                var details = matches.Count == 0
                    ? "(none)"
                    : string.Join("\n", matches.Select(m => $"{m.tag} | {m.hours}"));

                MessageBox.Show(
                    $"Matched rows: {matched}\n" +
                    $"Total hours: {sum:N2}\n\n" +
                    $"Raw matches:\n{details}",
                    "Matched Rows Detail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                var sampleTags = HoursImporter.Load(ofd.FileName)
                    .Select(r => r.Project)
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(12)
                    .ToList();

                var sampleText = sampleTags.Count > 0 ? string.Join("\n  - ", sampleTags) : "(none)";

                MessageBox.Show(
                    $"Hours file: {Path.GetFileName(ofd.FileName)}\n" +
                    $"Invoice Project Tag: \"{txtProjectTag.Text}\"\n\n" +
                    $"Matched rows: {matched}\n" +
                    $"Skipped rows: {skipped}\n" +
                    $"SUM(Column D for matches): {sum:N2}\n\n" +
                    $"Distinct Column B values seen:\n  - {sampleText}",
                    "Hours Import Audit (B filter / D sum)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                totalHours = HoursImporter.Load(ofd.FileName).Sum(r => r.Hours);
            }

            if (totalHours <= 0)
            {
                MessageBox.Show("No hours found for the current project in that file.", "Import Hours",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Remove any previous aggregate labor line
            foreach (var existing in _items
                .Where(i => i.ProductOrService.Equals("Service", StringComparison.OrdinalIgnoreCase)
                         && i.Description.StartsWith("Labor - All Employees", StringComparison.OrdinalIgnoreCase))
                .ToList())
            {
                _items.Remove(existing);
            }

            var desc = annotate ? $"Labor - All Employees ({totalHours:N2} hrs)" : "Labor - All Employees";
            _items.Add(new InvoiceItem
            {
                ProductOrService = "Service",
                Description = desc,
                Quantity = totalHours,
                Rate = nudDefaultLaborRate.Value
            });

            UpdateTotals();
        }

        // ---------- Totals ----------
        // MainForm.cs — FULL replacement of UpdateTotals()
        private void UpdateTotals()
        {
            // Do NOT clamp negatives; returns must stay negative.
            foreach (var it in _items)
            {
                if (string.IsNullOrWhiteSpace(it.ProductOrService))
                    it.ProductOrService = "Service";
                // keep whatever Quantity/Rate the importer produced (can be negative)
            }

            var materialsSubtotal = _items
                .Where(i => i.ProductOrService.Equals("Product", StringComparison.OrdinalIgnoreCase))
                .Sum(i => i.LineTotal);

            var servicesSubtotal = _items
                .Where(i => !i.ProductOrService.Equals("Product", StringComparison.OrdinalIgnoreCase))
                .Sum(i => i.LineTotal);

            var markup = Math.Round(materialsSubtotal * MaterialMarkupRate, 2);
            var baseSub = materialsSubtotal + servicesSubtotal;
            var subtotal = baseSub + markup;
            var tax = Math.Round(subtotal * (nudTaxRate.Value / 100m), 2);
            var total = subtotal + tax;

            lblSubtotal.Text = baseSub.ToString("C2", CultureInfo.CurrentCulture);
            lblMarkup.Text = markup.ToString("C2", CultureInfo.CurrentCulture);
            lblTax.Text = tax.ToString("C2", CultureInfo.CurrentCulture);
            lblTotal.Text = total.ToString("C2", CultureInfo.CurrentCulture);

            grid?.Refresh();
        }


    }
}
