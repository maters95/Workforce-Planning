import javax.swing.*;
import javax.swing.event.*;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;
import java.util.List;

/**
 * TestTerminal — a Swing CRT-style terminal that replicates the DRIVES
 * Batch Driving Record Request screens.
 *
 * Window title: "Test Terminal"
 * AppActivate("Test Terminal") in VBA will target this window.
 * keybd_event hardware keystrokes are dispatched normally by the JVM
 * (unlike Chromium, Java Swing processes OS-level WM_KEYDOWN/WM_KEYUP).
 *
 * Navigation the VBA macro performs (JAVA_TEST_MODE = True path):
 *   NavTo "3"  → Main Menu → Record Services
 *   NavTo "5"  → Batch Operations
 *   NavTo "14" → Batch Driving Record
 *   NavTo "7"  → Batch Driving Record Request (form, screen 4)
 *   TextTo "Y","N","Y"  → header fields (auto-advance after each char)
 *   KeyTo Enter         → go to Select TRAFFIC_REC_EMAIL (screen 5)
 *   KeyTo Tab           → focus Start List At Value input
 *   TextTo "S"          → type S into Start List field
 *   KeyTo F5            → return to form; focus hdr[0]
 *   KeyTo Tab           → hdr[0] → hdr[1]
 *   KeyTo Tab           → hdr[1] (email selected) → grid[0]
 *   TypeText / Tab loop → fill 90-cell grid
 *   KeyTo Enter         → submit; show completion message
 */
public class TestTerminal extends JFrame {

    // ── Colours ───────────────────────────────────────────────────
    private static final Color C_BG    = new Color(10, 10, 10);
    private static final Color C_GREEN = new Color(51, 255, 51);
    private static final Color C_CYAN  = new Color(85, 255, 255);
    private static final Color C_RED   = new Color(255, 85, 85);
    private static final Color C_DIM   = new Color(26, 140, 26);

    // ── Font ──────────────────────────────────────────────────────
    private static final Font MONO;
    static {
        Font mono = new Font("Courier New", Font.PLAIN, 13);
        for (String name : new String[]{"IBM Plex Mono", "Lucida Console", "Consolas"}) {
            Font t = new Font(name, Font.PLAIN, 13);
            if (t.getFamily().equalsIgnoreCase(name)) { mono = t; break; }
        }
        MONO = mono;
    }

    // ── Screen indices ────────────────────────────────────────────
    private static final int S_MAIN    = 0;
    private static final int S_RECSVC  = 1;
    private static final int S_BATCHOP = 2;
    private static final int S_BATCHDR = 3;
    private static final int S_FORM    = 4;
    private static final int S_EMAIL   = 5;

    // Menu screen config
    private static final String[] VALID_SEL = {"3", "5", "14", "7"};
    private static final int[]    MAX_OPT   = {6, 6, 15, 8};

    // ── State ─────────────────────────────────────────────────────
    private boolean emailSelected  = false;
    private boolean operationDone  = false;

    // ── UI components ─────────────────────────────────────────────
    private final CardLayout cardLayout = new CardLayout();
    private final JPanel     mainPanel  = new JPanel(cardLayout);

    private final JTextField[] menuInput = new JTextField[4];
    private final JLabel[]     menuError = new JLabel[4];

    private static final int ROWS     = 15;
    private static final int COLS     = 6;
    private static final int CELL_MAX = 8;

    private final JTextField[]   hdr        = new JTextField[3];
    private final JTextField[][] gridCell   = new JTextField[ROWS][COLS];
    private final List<JTextField> formOrder = new ArrayList<>();

    private JLabel emailDisplay;
    private JLabel statusLabel;

    private JTextField selStart;
    private JTextField selSel;

    // ── Constructor ───────────────────────────────────────────────
    public TestTerminal() {
        super("Test Terminal");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        getContentPane().setBackground(C_BG);
        setLayout(new BorderLayout());
        mainPanel.setBackground(C_BG);
        add(mainPanel, BorderLayout.CENTER);

        buildAllScreens();

        setPreferredSize(new Dimension(860, 640));
        pack();
        setLocationRelativeTo(null);
        setVisible(true);
        showScreen(S_MAIN);
    }

    // ═════════════════════════════════════════════════════════════
    //  Screen builders
    // ═════════════════════════════════════════════════════════════

    private void buildAllScreens() {
        buildMenu(0, "DRIVES - Main Menu", new String[]{
            "1.  System Administration",
            "2.  User Management",
            "3.  Record Services",
            "4.  Reporting & Analytics",
            "5.  Correspondence Management",
            "6.  System Utilities"
        });
        buildMenu(1, "Record Services", new String[]{
            "1.  Individual Record Lookup",
            "2.  Record Amendment",
            "3.  Record History",
            "4.  Print Services",
            "5.  Batch Operations",
            "6.  Record Archive"
        });
        buildMenu(2, "Batch Operations", new String[]{
            " 1.  Batch Licence Verification",
            " 2.  Batch Status Update",
            " 3.  Batch Print Request",
            " 4.  Batch Record Export",
            " 5.  Batch Record Import",
            " 6.  Batch Suspension Check",
            " 7.  Batch Address Update",
            " 8.  Batch Demerit Point Inquiry",
            " 9.  Batch Registration Renewal",
            "10.  Batch Penalty Notice",
            "11.  Batch Medical Review",
            "12.  Batch Interlock Audit",
            "13.  Batch Fee Assessment",
            "14.  Batch Driving Record",
            "15.  Batch Correspondence"
        });
        buildMenu(3, "Batch Driving Record", new String[]{
            "1.  Standard Record - Individual",
            "2.  Standard Record - Bulk Print",
            "3.  Traffic Section 12 - Individual",
            "4.  Traffic Section 12 - Bulk Print",
            "5.  Combined Record - Individual",
            "6.  Record Reprint Queue",
            "7.  Batch Driving Record Request",
            "8.  Batch Status Inquiry"
        });
        buildFormScreen();
        buildEmailSelectScreen();
    }

    private void buildMenu(final int idx, String title, String[] items) {
        JPanel panel = darkPanel(new BorderLayout());

        // Title
        JLabel ttl = makeLabel(title, C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(20, 0, 20, 0));
        panel.add(ttl, BorderLayout.NORTH);

        // Items
        JPanel itemsPanel = darkPanel(new GridLayout(items.length, 1, 0, 4));
        itemsPanel.setBorder(BorderFactory.createEmptyBorder(0, 80, 0, 80));
        for (String item : items) {
            itemsPanel.add(makeLabel(item, C_GREEN));
        }
        panel.add(itemsPanel, BorderLayout.CENTER);

        // Prompt area
        menuError[idx] = makeLabel("", C_RED);
        menuError[idx].setBorder(BorderFactory.createEmptyBorder(0, 24, 4, 0));

        JPanel inputRow = darkPanel(new FlowLayout(FlowLayout.LEFT, 24, 2));
        inputRow.add(makeLabel("Selection ==>", C_CYAN));
        JTextField inp = makeTextField(2);
        menuInput[idx] = inp;
        inputRow.add(inp);

        JPanel footer = darkPanel(new FlowLayout(FlowLayout.CENTER, 20, 2));
        for (String fk : new String[]{"F1 HELP", "F3 MENU", "F5 RETURN", "F6 CLEAR"}) {
            footer.add(makeLabel(fk, C_CYAN));
        }

        JPanel bottom = darkPanel(new BorderLayout());
        bottom.add(menuError[idx], BorderLayout.NORTH);
        bottom.add(inputRow, BorderLayout.CENTER);
        bottom.add(footer, BorderLayout.SOUTH);
        bottom.setBorder(BorderFactory.createEmptyBorder(8, 0, 8, 0));
        panel.add(bottom, BorderLayout.SOUTH);

        // Key bindings
        inp.setFocusTraversalKeysEnabled(false);
        inp.getInputMap(JComponent.WHEN_FOCUSED)
           .put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0), "enter");
        inp.getActionMap().put("enter", new AbstractAction() {
            public void actionPerformed(ActionEvent e) { handleMenuEnter(idx); }
        });
        inp.getInputMap(JComponent.WHEN_FOCUSED)
           .put(KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0), "tab");
        inp.getActionMap().put("tab", new AbstractAction() {
            public void actionPerformed(ActionEvent e) { /* ignore */ }
        });
        // F3/F5 go back one screen
        inp.getInputMap(JComponent.WHEN_FOCUSED)
           .put(KeyStroke.getKeyStroke(KeyEvent.VK_F3, 0), "back");
        inp.getInputMap(JComponent.WHEN_FOCUSED)
           .put(KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0), "back");
        inp.getActionMap().put("back", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                if (idx > 0) showScreen(idx - 1);
            }
        });

        mainPanel.add(panel, "screen" + idx);
    }

    private void handleMenuEnter(int idx) {
        String val = menuInput[idx].getText().trim();
        if (val.equals(VALID_SEL[idx])) {
            menuError[idx].setText("");
            if (idx == 3) resetBatchForm();
            showScreen(idx + 1);
        } else {
            try {
                int n = Integer.parseInt(val);
                if (n < 1 || n > MAX_OPT[idx]) {
                    menuError[idx].setText("Invalid selection. Enter a valid option number.");
                } else {
                    menuError[idx].setText("Function not available.");
                }
            } catch (NumberFormatException ex) {
                menuError[idx].setText("Invalid selection. Enter a valid option number.");
            }
            menuInput[idx].setText("");
        }
    }

    // ── Screen 4: Batch Driving Record Request ────────────────────
    private void buildFormScreen() {
        JPanel panel = darkPanel(new BorderLayout(0, 6));
        panel.setBorder(BorderFactory.createEmptyBorder(16, 20, 8, 20));

        JLabel ttl = makeLabel("Batch Driving Record Request", C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(0, 0, 14, 0));
        panel.add(ttl, BorderLayout.NORTH);

        JPanel body = darkPanel(new BorderLayout(0, 6));

        // Header row 1
        JPanel hrow1 = darkPanel(new FlowLayout(FlowLayout.LEFT, 8, 2));
        hrow1.add(makeLabel("Standard Driving Record Request:", C_CYAN));
        hdr[0] = makeTextField(1);
        hdr[0].setPreferredSize(new Dimension(18, 18));
        hrow1.add(hdr[0]);
        hrow1.add(Box.createHorizontalStrut(24));
        hrow1.add(makeLabel("Traffic Record Section 12 Request:", C_CYAN));
        hdr[1] = makeTextField(1);
        hdr[1].setPreferredSize(new Dimension(18, 18));
        hrow1.add(hdr[1]);

        // Header row 2
        JPanel hrow2 = darkPanel(new FlowLayout(FlowLayout.LEFT, 8, 2));
        hrow2.add(makeLabel("Email pdf:", C_CYAN));
        hdr[2] = makeTextField(1);
        hdr[2].setPreferredSize(new Dimension(18, 18));
        hrow2.add(hdr[2]);
        hrow2.add(Box.createHorizontalStrut(12));
        hrow2.add(makeLabel("Email Address:", C_CYAN));
        emailDisplay = makeLabel("", C_GREEN);
        hrow2.add(emailDisplay);

        JPanel hdrPanel = darkPanel(new BorderLayout(0, 4));
        hdrPanel.add(hrow1, BorderLayout.NORTH);
        hdrPanel.add(hrow2, BorderLayout.SOUTH);
        body.add(hdrPanel, BorderLayout.NORTH);

        // Grid: 15 rows × 6 cols
        JPanel gridPanel = darkPanel(new GridLayout(ROWS, COLS, 10, 4));
        gridPanel.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));
        for (int r = 0; r < ROWS; r++) {
            for (int c = 0; c < COLS; c++) {
                gridCell[r][c] = makeTextField(CELL_MAX);
                gridCell[r][c].setPreferredSize(new Dimension(84, 18));
                gridCell[r][c].setBorder(
                    BorderFactory.createMatteBorder(0, 0, 1, 0, C_DIM));
                gridPanel.add(gridCell[r][c]);
            }
        }
        body.add(gridPanel, BorderLayout.CENTER);

        // Footer + status
        JPanel foot = darkPanel(new BorderLayout());
        JPanel fkRow = darkPanel(new FlowLayout(FlowLayout.CENTER, 20, 2));
        for (String fk : new String[]{"F1 HELP", "F3 MENU", "F5 RETURN", "F6 CLEAR"}) {
            fkRow.add(makeLabel(fk, C_CYAN));
        }
        statusLabel = makeLabel(
            "The Operation has been successfully completed", C_CYAN);
        statusLabel.setHorizontalAlignment(SwingConstants.CENTER);
        statusLabel.setVisible(false);
        foot.add(fkRow, BorderLayout.CENTER);
        foot.add(statusLabel, BorderLayout.SOUTH);
        body.add(foot, BorderLayout.SOUTH);

        panel.add(body, BorderLayout.CENTER);

        // Build ordered list for focus traversal
        formOrder.clear();
        formOrder.add(hdr[0]);
        formOrder.add(hdr[1]);
        formOrder.add(hdr[2]);
        for (int r = 0; r < ROWS; r++) {
            for (int c = 0; c < COLS; c++) {
                formOrder.add(gridCell[r][c]);
            }
        }

        wireHeaderInputs();
        wireGridInputs();

        mainPanel.add(panel, "screen4");
    }

    private void wireHeaderInputs() {
        for (int hi = 0; hi < 3; hi++) {
            final int h = hi;
            final JTextField tf = hdr[hi];
            tf.setFocusTraversalKeysEnabled(false);

            // Auto-uppercase + auto-advance for hdr[0] and hdr[1]
            tf.getDocument().addDocumentListener(new SimpleDocListener() {
                public void onInsert() {
                    SwingUtilities.invokeLater(new Runnable() {
                        public void run() {
                            String v = tf.getText().toUpperCase();
                            if (!tf.getText().equals(v)) {
                                tf.setText(v);
                                tf.setCaretPosition(v.length());
                            }
                            if (h < 2 && tf.getText().length() >= 1) {
                                focusFormNext(tf);
                            }
                        }
                    });
                }
            });

            // Tab
            tf.getInputMap(JComponent.WHEN_FOCUSED)
              .put(KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0), "tab");
            tf.getActionMap().put("tab", new AbstractAction() {
                public void actionPerformed(ActionEvent e) {
                    if (operationDone) return;
                    if ((h == 1 || h == 2) && emailSelected) {
                        // Skip straight to grid
                        formOrder.get(3).requestFocusInWindow();
                    } else {
                        focusFormNext(tf);
                    }
                }
            });

            // Enter → go to email select
            tf.getInputMap(JComponent.WHEN_FOCUSED)
              .put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0), "enter");
            tf.getActionMap().put("enter", new AbstractAction() {
                public void actionPerformed(ActionEvent e) {
                    if (!operationDone && !emailSelected) showScreen(S_EMAIL);
                }
            });

            // F3/F5 → back
            tf.getInputMap(JComponent.WHEN_FOCUSED)
              .put(KeyStroke.getKeyStroke(KeyEvent.VK_F3, 0), "back");
            tf.getInputMap(JComponent.WHEN_FOCUSED)
              .put(KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0), "back");
            tf.getActionMap().put("back", new AbstractAction() {
                public void actionPerformed(ActionEvent e) { goBack(); }
            });
        }
    }

    private void wireGridInputs() {
        for (int r = 0; r < ROWS; r++) {
            for (int c = 0; c < COLS; c++) {
                final JTextField tf = gridCell[r][c];
                tf.setFocusTraversalKeysEnabled(false);

                // Auto-uppercase + auto-advance at max length
                tf.getDocument().addDocumentListener(new SimpleDocListener() {
                    public void onInsert() {
                        SwingUtilities.invokeLater(new Runnable() {
                            public void run() {
                                String v = tf.getText().toUpperCase();
                                if (!tf.getText().equals(v)) {
                                    tf.setText(v);
                                    tf.setCaretPosition(v.length());
                                }
                                if (tf.getText().length() >= CELL_MAX) {
                                    focusFormNext(tf);
                                }
                            }
                        });
                    }
                });

                // Focus highlight (dim underline → bright underline)
                tf.addFocusListener(new FocusAdapter() {
                    public void focusGained(FocusEvent e) {
                        tf.setBorder(BorderFactory.createMatteBorder(
                            0, 0, 1, 0, C_GREEN));
                    }
                    public void focusLost(FocusEvent e) {
                        tf.setBorder(BorderFactory.createMatteBorder(
                            0, 0, 1, 0, C_DIM));
                    }
                });

                // Tab
                tf.getInputMap(JComponent.WHEN_FOCUSED)
                  .put(KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0), "tab");
                tf.getActionMap().put("tab", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        if (!operationDone) focusFormNext(tf);
                    }
                });

                // Enter → submit if any grid cell has data
                tf.getInputMap(JComponent.WHEN_FOCUSED)
                  .put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0), "enter");
                tf.getActionMap().put("enter", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        if (operationDone) return;
                        boolean hasData = false;
                        for (JTextField f : formOrder) {
                            if (!f.getText().trim().isEmpty()) {
                                hasData = true;
                                break;
                            }
                        }
                        if (hasData) {
                            operationDone = true;
                            statusLabel.setVisible(true);
                        }
                    }
                });

                // F3/F5 → back
                tf.getInputMap(JComponent.WHEN_FOCUSED)
                  .put(KeyStroke.getKeyStroke(KeyEvent.VK_F3, 0), "back");
                tf.getInputMap(JComponent.WHEN_FOCUSED)
                  .put(KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0), "back");
                tf.getActionMap().put("back", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) { goBack(); }
                });
            }
        }
    }

    // ── Screen 5: Select TRAFFIC_REC_EMAIL ────────────────────────
    private void buildEmailSelectScreen() {
        JPanel panel = darkPanel(new BorderLayout(0, 0));
        panel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        JLabel ttl = makeLabel("Select TRAFFIC_REC_EMAIL", C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(0, 0, 20, 0));
        panel.add(ttl, BorderLayout.NORTH);

        JPanel body = darkPanel(new BorderLayout(0, 10));
        body.setBorder(BorderFactory.createEmptyBorder(0, 20, 0, 20));

        // Start row (right-aligned)
        JPanel startRow = darkPanel(new FlowLayout(FlowLayout.RIGHT, 8, 2));
        startRow.add(makeLabel("Start List At Value:", C_CYAN));
        selStart = makeTextField(30);
        selStart.setPreferredSize(new Dimension(220, 20));
        startRow.add(selStart);
        body.add(startRow, BorderLayout.NORTH);

        // Table header
        JPanel tableArea = darkPanel(new BorderLayout(0, 6));
        JPanel tableHead = darkPanel(new FlowLayout(FlowLayout.LEFT, 16, 2));
        JLabel selHdr = makeLabel("Sel", C_CYAN);
        selHdr.setFont(MONO.deriveFont(Font.BOLD));
        tableHead.add(selHdr);
        JLabel valHdr = makeLabel("Value", C_CYAN);
        valHdr.setFont(MONO.deriveFont(Font.BOLD));
        tableHead.add(valHdr);

        // Single email row
        JPanel emailRow = darkPanel(new FlowLayout(FlowLayout.LEFT, 16, 2));
        selSel = makeTextField(1);
        selSel.setPreferredSize(new Dimension(18, 18));
        emailRow.add(selSel);
        emailRow.add(makeLabel("CAU", C_GREEN));
        emailRow.add(Box.createHorizontalStrut(24));
        emailRow.add(makeLabel("batchemail@mail.com", C_GREEN));

        tableArea.add(tableHead, BorderLayout.NORTH);
        tableArea.add(emailRow, BorderLayout.CENTER);
        body.add(tableArea, BorderLayout.CENTER);
        panel.add(body, BorderLayout.CENTER);

        // Footer
        JPanel foot = darkPanel(new FlowLayout(FlowLayout.CENTER, 20, 4));
        for (String fk : new String[]{"F1 Help", "F5 Return", "F7 Prev", "F8 Next"}) {
            foot.add(makeLabel(fk, C_CYAN));
        }
        panel.add(foot, BorderLayout.SOUTH);

        // Wire selStart
        selStart.setFocusTraversalKeysEnabled(false);
        selStart.getInputMap(JComponent.WHEN_FOCUSED)
                .put(KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0), "tab");
        selStart.getActionMap().put("tab", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                selSel.requestFocusInWindow();
            }
        });
        bindF5toEmailReturn(selStart);

        // Wire selSel
        selSel.setFocusTraversalKeysEnabled(false);
        selSel.getDocument().addDocumentListener(new SimpleDocListener() {
            public void onInsert() {
                SwingUtilities.invokeLater(new Runnable() {
                    public void run() {
                        String v = selSel.getText().toUpperCase();
                        if (!selSel.getText().equals(v)) selSel.setText(v);
                    }
                });
            }
        });
        selSel.getInputMap(JComponent.WHEN_FOCUSED)
              .put(KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0), "tab");
        selSel.getActionMap().put("tab", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                selStart.requestFocusInWindow();
            }
        });
        bindF5toEmailReturn(selSel);

        mainPanel.add(panel, "screen5");
    }

    private void bindF5toEmailReturn(JTextField tf) {
        tf.getInputMap(JComponent.WHEN_FOCUSED)
          .put(KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0), "f5");
        tf.getActionMap().put("f5", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                emailSelected = true;
                emailDisplay.setText("batchemail@mail.com");
                showScreen(S_FORM);
            }
        });
    }

    // ═════════════════════════════════════════════════════════════
    //  Navigation helpers
    // ═════════════════════════════════════════════════════════════

    private void showScreen(final int idx) {
        cardLayout.show(mainPanel, "screen" + idx);
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                switch (idx) {
                    case S_MAIN:
                    case S_RECSVC:
                    case S_BATCHOP:
                    case S_BATCHDR:
                        menuInput[idx].setText("");
                        menuError[idx].setText("");
                        menuInput[idx].requestFocusInWindow();
                        break;
                    case S_FORM:
                        hdr[0].requestFocusInWindow();
                        break;
                    case S_EMAIL:
                        selStart.setText("");
                        selSel.setText("");
                        selStart.requestFocusInWindow();
                        break;
                }
            }
        });
    }

    private void resetBatchForm() {
        for (JTextField tf : hdr) tf.setText("");
        for (int r = 0; r < ROWS; r++) {
            for (int c = 0; c < COLS; c++) {
                gridCell[r][c].setText("");
                gridCell[r][c].setBorder(
                    BorderFactory.createMatteBorder(0, 0, 1, 0, C_DIM));
            }
        }
        emailSelected = false;
        emailDisplay.setText("");
        operationDone = false;
        statusLabel.setVisible(false);
    }

    private void goBack() {
        resetBatchForm();
        showScreen(S_BATCHDR);
    }

    private void focusFormNext(JTextField tf) {
        int i = formOrder.indexOf(tf);
        if (i >= 0 && i < formOrder.size() - 1) {
            formOrder.get(i + 1).requestFocusInWindow();
        }
    }

    // ═════════════════════════════════════════════════════════════
    //  UI helpers
    // ═════════════════════════════════════════════════════════════

    private JPanel darkPanel(LayoutManager lm) {
        JPanel p = new JPanel(lm);
        p.setBackground(C_BG);
        return p;
    }

    private JLabel makeLabel(String text, Color color) {
        JLabel l = new JLabel(text);
        l.setForeground(color);
        l.setFont(MONO);
        l.setBackground(C_BG);
        l.setOpaque(true);
        return l;
    }

    private JTextField makeTextField(int maxLen) {
        JTextField tf = new JTextField();
        tf.setBackground(C_BG);
        tf.setForeground(C_GREEN);
        tf.setCaretColor(C_GREEN);
        tf.setFont(MONO);
        tf.setBorder(BorderFactory.createEmptyBorder(0, 2, 0, 2));
        tf.setOpaque(true);
        tf.setSelectionColor(C_DIM);
        tf.setSelectedTextColor(C_GREEN);
        ((AbstractDocument) tf.getDocument())
            .setDocumentFilter(new LengthFilter(maxLen));
        return tf;
    }

    // ═════════════════════════════════════════════════════════════
    //  Inner classes
    // ═════════════════════════════════════════════════════════════

    /** DocumentFilter that caps input at a maximum character count. */
    private static class LengthFilter extends DocumentFilter {
        private final int max;
        LengthFilter(int max) { this.max = max; }

        @Override
        public void insertString(FilterBypass fb, int offset,
                String str, AttributeSet attr) throws BadLocationException {
            if (str == null) return;
            int avail = max - fb.getDocument().getLength();
            if (avail <= 0) return;
            if (str.length() > avail) str = str.substring(0, avail);
            super.insertString(fb, offset, str, attr);
        }

        @Override
        public void replace(FilterBypass fb, int offset, int length,
                String str, AttributeSet attr) throws BadLocationException {
            if (str == null) str = "";
            int remain = fb.getDocument().getLength() - length;
            int avail  = max - remain;
            if (avail <= 0 && !str.isEmpty()) return;
            if (str.length() > avail) str = str.substring(0, avail);
            super.replace(fb, offset, length, str, attr);
        }
    }

    /** Convenience DocumentListener that only requires implementing onInsert(). */
    private abstract static class SimpleDocListener implements DocumentListener {
        public abstract void onInsert();
        public void insertUpdate(DocumentEvent e)  { onInsert(); }
        public void removeUpdate(DocumentEvent e)  {}
        public void changedUpdate(DocumentEvent e) {}
    }

    // ═════════════════════════════════════════════════════════════
    //  Main
    // ═════════════════════════════════════════════════════════════

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() { new TestTerminal(); }
        });
    }
}
