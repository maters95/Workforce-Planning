import javax.swing.*;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;
import java.util.List;

/**
 * TestTerminal — Swing CRT-style terminal that replicates the DRIVES
 * Batch Driving Record Request screens.
 *
 * Window title: "Test Terminal"
 * AppActivate("Test Terminal") in VBA targets this window.
 * keybd_event hardware keystrokes work in Swing (unlike Chromium).
 *
 * Build:  javac TestTerminal.java
 * Run:    java TestTerminal
 *
 * VBA keystroke sequence (JAVA_TEST_MODE = True path):
 *   NavTo 3,5,14,7        menu navigation to batch form (screen 4)
 *   TextTo Y,N,Y          header fields (auto-advance after each char)
 *   KeyTo Enter           → go to TRAFFIC_REC_EMAIL screen (screen 5)
 *   KeyTo Tab             selStart → selSel
 *   TextTo S              type into selSel
 *   KeyTo F5              return to form; focus hdr[0]
 *   KeyTo Tab × 2         hdr[0]→hdr[1] → (emailSelected) → grid[0]
 *   TypeText+Tab loop     fill grid (Tab for <8-char nums; auto-advance at 8)
 *   KeyTo Enter           submit; show completion status
 *   (overflow) KeyTo F5   → screen 3; repeat navigation
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
        Font f = new Font("Courier New", Font.PLAIN, 13);
        for (String n : new String[]{"IBM Plex Mono","Lucida Console","Consolas"}) {
            Font t = new Font(n, Font.PLAIN, 13);
            if (t.getFamily().equalsIgnoreCase(n)) { f = t; break; }
        }
        MONO = f;
    }

    // ── Screen indices ────────────────────────────────────────────
    private static final int S_MAIN = 0, S_REC = 1, S_BATCH = 2,
                              S_BDR  = 3, S_FORM = 4, S_EMAIL = 5;

    private static final String[] VALID = {"3","5","14","7"};
    private static final int[]    MAXOP = {6, 6, 15, 8};

    // ── State ─────────────────────────────────────────────────────
    private boolean emailSelected = false;
    private boolean operationDone = false;

    // ── UI ────────────────────────────────────────────────────────
    private final CardLayout cardLayout = new CardLayout();
    private final JPanel     cards      = new JPanel(cardLayout);

    private final JTextField[] menuInput = new JTextField[4];
    private final JLabel[]     menuError = new JLabel[4];

    private static final int ROWS = 15, COLS = 6, CELL_MAX = 8;

    private final JTextField[]   hdr      = new JTextField[3];
    private final JTextField[][] grid     = new JTextField[ROWS][COLS];
    private final List<JTextField> order  = new ArrayList<>();

    private JLabel emailDisplay;
    private JLabel statusLabel;
    private JTextField selStart;
    private JTextField selSel;

    // ─────────────────────────────────────────────────────────────
    public TestTerminal() {
        super("Test Terminal");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setResizable(false);
        setAlwaysOnTop(true);
        getContentPane().setBackground(C_BG);
        setLayout(new BorderLayout());
        cards.setBackground(C_BG);
        add(cards, BorderLayout.CENTER);

        buildMenus();
        buildFormScreen();
        buildEmailScreen();

        setPreferredSize(new Dimension(900, 660));
        pack();
        setLocationRelativeTo(null);
        setVisible(true);
        showScreen(S_MAIN);
    }

    // ═════════════════════════════════════════════════════════════
    //  Menu screens 0-3
    // ═════════════════════════════════════════════════════════════

    private void buildMenus() {
        String[][] items = {
            {"1.  System Administration","2.  User Management",
             "3.  Record Services","4.  Reporting & Analytics",
             "5.  Correspondence Management","6.  System Utilities"},
            {"1.  Individual Record Lookup","2.  Record Amendment",
             "3.  Record History","4.  Print Services",
             "5.  Batch Operations","6.  Record Archive"},
            {" 1.  Batch Licence Verification"," 2.  Batch Status Update",
             " 3.  Batch Print Request"," 4.  Batch Record Export",
             " 5.  Batch Record Import"," 6.  Batch Suspension Check",
             " 7.  Batch Address Update"," 8.  Batch Demerit Point Inquiry",
             " 9.  Batch Registration Renewal","10.  Batch Penalty Notice",
             "11.  Batch Medical Review","12.  Batch Interlock Audit",
             "13.  Batch Fee Assessment","14.  Batch Driving Record",
             "15.  Batch Correspondence"},
            {"1.  Standard Record - Individual","2.  Standard Record - Bulk Print",
             "3.  Traffic Section 12 - Individual","4.  Traffic Section 12 - Bulk Print",
             "5.  Combined Record - Individual","6.  Record Reprint Queue",
             "7.  Batch Driving Record Request","8.  Batch Status Inquiry"}
        };
        String[] titles = {
            "DRIVES - Main Menu","Record Services",
            "Batch Operations","Batch Driving Record"
        };
        for (int i = 0; i < 4; i++) buildMenu(i, titles[i], items[i]);
    }

    private void buildMenu(final int idx, String title, String[] items) {
        JPanel p = dark(new BorderLayout());

        JLabel ttl = lbl(title, C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(20,0,18,0));
        p.add(ttl, BorderLayout.NORTH);

        JPanel list = dark(new GridLayout(items.length, 1, 0, 5));
        list.setBorder(BorderFactory.createEmptyBorder(0, 80, 0, 80));
        for (String it : items) list.add(lbl(it, C_GREEN));
        p.add(list, BorderLayout.CENTER);

        menuError[idx] = lbl("", C_RED);
        menuError[idx].setBorder(BorderFactory.createEmptyBorder(0,24,3,0));

        JPanel inputRow = dark(new FlowLayout(FlowLayout.LEFT, 24, 2));
        inputRow.add(lbl("Selection ==>", C_CYAN));
        JTextField inp = field(2);
        menuInput[idx] = inp;
        inputRow.add(inp);

        JPanel fkRow = dark(new FlowLayout(FlowLayout.CENTER, 20, 2));
        for (String fk : new String[]{"F1 HELP","F3 MENU","F5 RETURN","F6 CLEAR"})
            fkRow.add(lbl(fk, C_CYAN));

        JPanel bot = dark(new BorderLayout());
        bot.add(menuError[idx], BorderLayout.NORTH);
        bot.add(inputRow,       BorderLayout.CENTER);
        bot.add(fkRow,          BorderLayout.SOUTH);
        bot.setBorder(BorderFactory.createEmptyBorder(6,0,8,0));
        p.add(bot, BorderLayout.SOUTH);

        // Enter
        inp.setFocusTraversalKeysEnabled(false);
        bind(inp, KeyEvent.VK_ENTER, 0, new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                String v = menuInput[idx].getText().trim();
                if (v.equals(VALID[idx])) {
                    menuError[idx].setText("");
                    if (idx == 3) resetForm();
                    showScreen(idx + 1);
                } else {
                    try {
                        int n = Integer.parseInt(v);
                        menuError[idx].setText(
                            n < 1 || n > MAXOP[idx]
                            ? "Invalid selection. Enter a valid option number."
                            : "Function not available.");
                    } catch (NumberFormatException ex) {
                        menuError[idx].setText(
                            "Invalid selection. Enter a valid option number.");
                    }
                    menuInput[idx].setText("");
                }
            }
        });
        // F3/F5 → back
        AbstractAction back = new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                if (idx > 0) showScreen(idx - 1);
            }
        };
        bind(inp, KeyEvent.VK_F3, 0, back);
        bind(inp, KeyEvent.VK_F5, 0, back);
        // Tab → swallow
        bind(inp, KeyEvent.VK_TAB, 0, new AbstractAction() {
            public void actionPerformed(ActionEvent e) {}
        });

        cards.add(p, "s" + idx);
    }

    // ═════════════════════════════════════════════════════════════
    //  Screen 4: Batch Driving Record Request
    // ═════════════════════════════════════════════════════════════

    private void buildFormScreen() {
        JPanel p = dark(new BorderLayout(0, 4));
        p.setBorder(BorderFactory.createEmptyBorder(14, 18, 6, 18));

        JLabel ttl = lbl("Batch Driving Record Request", C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(0,0,12,0));
        p.add(ttl, BorderLayout.NORTH);

        // Header row 1
        hdr[0] = field(1); hdr[1] = field(1); hdr[2] = field(1);
        for (JTextField h : hdr) h.setPreferredSize(new Dimension(18, 18));

        JPanel hr1 = dark(new FlowLayout(FlowLayout.LEFT, 6, 2));
        hr1.add(lbl("Standard Driving Record Request:", C_CYAN)); hr1.add(hdr[0]);
        hr1.add(Box.createHorizontalStrut(16));
        hr1.add(lbl("Traffic Record Section 12 Request:", C_CYAN)); hr1.add(hdr[1]);

        JPanel hr2 = dark(new FlowLayout(FlowLayout.LEFT, 6, 2));
        hr2.add(lbl("Email pdf:", C_CYAN)); hr2.add(hdr[2]);
        hr2.add(Box.createHorizontalStrut(10));
        hr2.add(lbl("Email Address:", C_CYAN));
        emailDisplay = lbl("", C_GREEN);
        hr2.add(emailDisplay);

        JPanel hdrPanel = dark(new BorderLayout(0, 3));
        hdrPanel.add(hr1, BorderLayout.NORTH);
        hdrPanel.add(hr2, BorderLayout.SOUTH);

        // Grid
        JPanel gridPanel = dark(new GridLayout(ROWS, COLS, 8, 4));
        gridPanel.setBorder(BorderFactory.createEmptyBorder(8,16,8,16));
        for (int r = 0; r < ROWS; r++)
            for (int c = 0; c < COLS; c++) {
                grid[r][c] = field(CELL_MAX);
                grid[r][c].setPreferredSize(new Dimension(84, 18));
                grid[r][c].setBorder(
                    BorderFactory.createMatteBorder(0,0,1,0, C_DIM));
                gridPanel.add(grid[r][c]);
            }

        JPanel body = dark(new BorderLayout(0, 4));
        body.add(hdrPanel,  BorderLayout.NORTH);
        body.add(gridPanel, BorderLayout.CENTER);

        // Footer + status
        JPanel fkRow = dark(new FlowLayout(FlowLayout.CENTER, 20, 2));
        for (String fk : new String[]{"F1 HELP","F3 MENU","F5 RETURN","F6 CLEAR"})
            fkRow.add(lbl(fk, C_CYAN));
        statusLabel = lbl("The Operation has been successfully completed", C_CYAN);
        statusLabel.setHorizontalAlignment(SwingConstants.CENTER);
        statusLabel.setVisible(false);
        JPanel bot = dark(new BorderLayout());
        bot.add(fkRow,       BorderLayout.CENTER);
        bot.add(statusLabel, BorderLayout.SOUTH);
        body.add(bot, BorderLayout.SOUTH);

        p.add(body, BorderLayout.CENTER);

        // Build ordered traversal list
        order.clear();
        order.add(hdr[0]); order.add(hdr[1]); order.add(hdr[2]);
        for (int r = 0; r < ROWS; r++)
            for (int c = 0; c < COLS; c++)
                order.add(grid[r][c]);

        wireHeaders();
        wireGrid();

        cards.add(p, "s4");
    }

    private void wireHeaders() {
        for (int hi = 0; hi < 3; hi++) {
            final int   h  = hi;
            final JTextField tf = hdr[hi];
            tf.setFocusTraversalKeysEnabled(false);

            // Auto-uppercase + auto-advance (hdr[0] and hdr[1] only)
            tf.getDocument().addDocumentListener(new DocIns() {
                public void onInsert() {
                    SwingUtilities.invokeLater(new Runnable() { public void run() {
                        if (operationDone) return;
                        upperCase(tf);
                        if (h < 2 && tf.getText().length() >= 1) next(tf);
                    }});
                }
            });

            // Tab
            bind(tf, KeyEvent.VK_TAB, 0, new AbstractAction() {
                public void actionPerformed(ActionEvent e) {
                    if (operationDone) return;
                    // hdr[1] or hdr[2] after email selected → skip to grid
                    if (h >= 1 && emailSelected) order.get(3).requestFocusInWindow();
                    else next(tf);
                }
            });

            // Enter → email screen (only if email not yet selected)
            bind(tf, KeyEvent.VK_ENTER, 0, new AbstractAction() {
                public void actionPerformed(ActionEvent e) {
                    if (operationDone) return;
                    if (!emailSelected) showScreen(S_EMAIL);
                }
            });

            // F3/F5 → back to screen 3
            AbstractAction back = new AbstractAction() {
                public void actionPerformed(ActionEvent e) { goBack(); }
            };
            bind(tf, KeyEvent.VK_F3, 0, back);
            bind(tf, KeyEvent.VK_F5, 0, back);
        }
    }

    private void wireGrid() {
        for (int r = 0; r < ROWS; r++)
            for (int c = 0; c < COLS; c++) {
                final JTextField tf = grid[r][c];
                tf.setFocusTraversalKeysEnabled(false);

                // Underline highlight on focus
                tf.addFocusListener(new FocusAdapter() {
                    public void focusGained(FocusEvent e) {
                        tf.setBorder(BorderFactory.createMatteBorder(0,0,1,0,C_GREEN));
                    }
                    public void focusLost(FocusEvent e) {
                        tf.setBorder(BorderFactory.createMatteBorder(0,0,1,0,C_DIM));
                    }
                });

                // Auto-uppercase + auto-advance at CELL_MAX
                tf.getDocument().addDocumentListener(new DocIns() {
                    public void onInsert() {
                        SwingUtilities.invokeLater(new Runnable() { public void run() {
                            if (operationDone) return;
                            upperCase(tf);
                            if (tf.getText().length() >= CELL_MAX) next(tf);
                        }});
                    }
                });

                // Tab → next cell
                bind(tf, KeyEvent.VK_TAB, 0, new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        if (!operationDone) next(tf);
                    }
                });

                // Enter → submit
                bind(tf, KeyEvent.VK_ENTER, 0, new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        if (operationDone) return;
                        for (JTextField f : order)
                            if (!f.getText().trim().isEmpty()) {
                                operationDone = true;
                                statusLabel.setVisible(true);
                                return;
                            }
                    }
                });

                // F3/F5 → back
                AbstractAction back = new AbstractAction() {
                    public void actionPerformed(ActionEvent e) { goBack(); }
                };
                bind(tf, KeyEvent.VK_F3, 0, back);
                bind(tf, KeyEvent.VK_F5, 0, back);
            }
    }

    // ═════════════════════════════════════════════════════════════
    //  Screen 5: Select TRAFFIC_REC_EMAIL
    // ═════════════════════════════════════════════════════════════

    private void buildEmailScreen() {
        JPanel p = dark(new BorderLayout(0,0));
        p.setBorder(BorderFactory.createEmptyBorder(20,20,20,20));

        JLabel ttl = lbl("Select TRAFFIC_REC_EMAIL", C_CYAN);
        ttl.setFont(MONO.deriveFont(Font.BOLD, 14f));
        ttl.setHorizontalAlignment(SwingConstants.CENTER);
        ttl.setBorder(BorderFactory.createEmptyBorder(0,0,20,0));
        p.add(ttl, BorderLayout.NORTH);

        JPanel body = dark(new BorderLayout(0,10));
        body.setBorder(BorderFactory.createEmptyBorder(0,20,0,20));

        // Start List At Value (right-aligned)
        JPanel startRow = dark(new FlowLayout(FlowLayout.RIGHT, 8, 2));
        startRow.add(lbl("Start List At Value:", C_CYAN));
        selStart = field(30);
        selStart.setPreferredSize(new Dimension(220, 20));
        selStart.setBorder(BorderFactory.createMatteBorder(0,0,1,0,C_DIM));
        startRow.add(selStart);
        body.add(startRow, BorderLayout.NORTH);

        // Column header
        JPanel colHead = dark(new FlowLayout(FlowLayout.LEFT, 16, 2));
        JLabel sh = lbl("Sel", C_CYAN); sh.setFont(MONO.deriveFont(Font.BOLD));
        JLabel vh = lbl("Value", C_CYAN); vh.setFont(MONO.deriveFont(Font.BOLD));
        colHead.add(sh); colHead.add(vh);

        // Single row: sel input | CAU | batchemail@mail.com
        JPanel emailRow = dark(new FlowLayout(FlowLayout.LEFT, 16, 4));
        selSel = field(1);
        selSel.setPreferredSize(new Dimension(18, 18));
        emailRow.add(selSel);
        emailRow.add(lbl("CAU", C_GREEN));
        emailRow.add(Box.createHorizontalStrut(24));
        emailRow.add(lbl("batchemail@mail.com", C_GREEN));

        JPanel tbl = dark(new BorderLayout(0,4));
        tbl.add(colHead,  BorderLayout.NORTH);
        tbl.add(emailRow, BorderLayout.CENTER);
        body.add(tbl, BorderLayout.CENTER);
        p.add(body, BorderLayout.CENTER);

        JPanel fkRow = dark(new FlowLayout(FlowLayout.CENTER, 20, 4));
        for (String fk : new String[]{"F1 Help","F5 Return","F7 Prev","F8 Next"})
            fkRow.add(lbl(fk, C_CYAN));
        p.add(fkRow, BorderLayout.SOUTH);

        // Wire selStart: Tab → selSel, F5 → return to form
        selStart.setFocusTraversalKeysEnabled(false);
        bind(selStart, KeyEvent.VK_TAB, 0, new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                selSel.requestFocusInWindow();
            }
        });
        bindF5Return(selStart);

        // Wire selSel: Tab → eat (forward Tab does nothing, like HTML), F5 → return
        selSel.setFocusTraversalKeysEnabled(false);
        selSel.getDocument().addDocumentListener(new DocIns() {
            public void onInsert() {
                SwingUtilities.invokeLater(new Runnable() { public void run() {
                    upperCase(selSel);
                }});
            }
        });
        bind(selSel, KeyEvent.VK_TAB, 0, new AbstractAction() {
            // Forward Tab eats keystroke (matches HTML: no-op forward Tab on selInput)
            public void actionPerformed(ActionEvent e) {}
        });
        bindF5Return(selSel);

        // Eat unused F-keys on both
        for (JTextField tf : new JTextField[]{selStart, selSel})
            for (int vk : new int[]{KeyEvent.VK_F1, KeyEvent.VK_F3,
                                     KeyEvent.VK_F7, KeyEvent.VK_F8})
                bind(tf, vk, 0, new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {}
                });

        cards.add(p, "s5");
    }

    private void bindF5Return(JTextField tf) {
        bind(tf, KeyEvent.VK_F5, 0, new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                emailSelected = true;
                emailDisplay.setText("batchemail@mail.com");
                showScreen(S_FORM);
            }
        });
    }

    // ═════════════════════════════════════════════════════════════
    //  Navigation
    // ═════════════════════════════════════════════════════════════

    private void showScreen(final int idx) {
        cardLayout.show(cards, "s" + idx);
        SwingUtilities.invokeLater(new Runnable() { public void run() {
            switch (idx) {
                case S_MAIN: case S_REC: case S_BATCH: case S_BDR:
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
        }});
    }

    private void resetForm() {
        for (JTextField h : hdr) h.setText("");
        for (int r = 0; r < ROWS; r++)
            for (int c = 0; c < COLS; c++) {
                grid[r][c].setText("");
                grid[r][c].setBorder(
                    BorderFactory.createMatteBorder(0,0,1,0,C_DIM));
            }
        emailSelected = false;
        emailDisplay.setText("");
        operationDone = false;
        statusLabel.setVisible(false);
    }

    private void goBack() {
        resetForm();
        showScreen(S_BDR);
    }

    private void next(JTextField tf) {
        int i = order.indexOf(tf);
        if (i >= 0 && i < order.size() - 1) order.get(i+1).requestFocusInWindow();
    }

    // ═════════════════════════════════════════════════════════════
    //  Helpers
    // ═════════════════════════════════════════════════════════════

    private JPanel dark(LayoutManager lm) {
        JPanel p = new JPanel(lm);
        p.setBackground(C_BG);
        return p;
    }

    private JLabel lbl(String text, Color color) {
        JLabel l = new JLabel(text);
        l.setForeground(color); l.setFont(MONO);
        l.setBackground(C_BG); l.setOpaque(true);
        return l;
    }

    /**
     * Creates a styled text field with:
     *  - Green-on-black CRT appearance
     *  - Length cap via DocumentFilter
     *  - Paste/cut/drag disabled (no injecting arbitrary text)
     *  - No select-all on focus (caret moves to end instead)
     */
    private JTextField field(int maxLen) {
        JTextField tf = new JTextField() {
            // Prevent default select-all that Swing does on some focus events
            public void selectAll() {}
        };
        tf.setBackground(C_BG);
        tf.setForeground(C_GREEN);
        tf.setCaretColor(C_GREEN);
        tf.setFont(MONO);
        tf.setBorder(BorderFactory.createEmptyBorder(0,2,0,2));
        tf.setOpaque(true);
        tf.setSelectionColor(C_DIM);
        tf.setSelectedTextColor(C_GREEN);
        ((AbstractDocument) tf.getDocument()).setDocumentFilter(new MaxLen(maxLen));

        // Caret to end on focus (deselect any accidental selection)
        tf.addFocusListener(new FocusAdapter() {
            public void focusGained(FocusEvent e) {
                SwingUtilities.invokeLater(new Runnable() { public void run() {
                    tf.setCaretPosition(tf.getText().length());
                }});
            }
        });

        // Block Ctrl+V (paste), Ctrl+X (cut)
        int ctrl = InputEvent.CTRL_DOWN_MASK;
        tf.getInputMap(JComponent.WHEN_FOCUSED)
          .put(KeyStroke.getKeyStroke(KeyEvent.VK_V, ctrl), "none");
        tf.getInputMap(JComponent.WHEN_FOCUSED)
          .put(KeyStroke.getKeyStroke(KeyEvent.VK_X, ctrl), "none");
        // Block drag-drop
        tf.setTransferHandler(null);

        return tf;
    }

    private static void upperCase(JTextField tf) {
        String v = tf.getText().toUpperCase();
        if (!tf.getText().equals(v)) {
            int pos = tf.getCaretPosition();
            tf.setText(v);
            tf.setCaretPosition(Math.min(pos, v.length()));
        }
    }

    private static void bind(JTextField tf, int vk, int mod, Action a) {
        String key = "k_" + vk + "_" + mod;
        tf.getInputMap(JComponent.WHEN_FOCUSED)
          .put(KeyStroke.getKeyStroke(vk, mod), key);
        tf.getActionMap().put(key, a);
    }

    // ═════════════════════════════════════════════════════════════
    //  Inner helpers
    // ═════════════════════════════════════════════════════════════

    private static class MaxLen extends DocumentFilter {
        private final int max;
        MaxLen(int max) { this.max = max; }
        public void insertString(FilterBypass fb, int off, String s, AttributeSet a)
                throws BadLocationException {
            if (s == null) return;
            int avail = max - fb.getDocument().getLength();
            if (avail > 0) super.insertString(fb, off,
                s.substring(0, Math.min(s.length(), avail)), a);
        }
        public void replace(FilterBypass fb, int off, int len, String s, AttributeSet a)
                throws BadLocationException {
            if (s == null) s = "";
            int remain = fb.getDocument().getLength() - len;
            int avail  = max - remain;
            if (avail <= 0 && s.isEmpty()) { super.replace(fb, off, len, s, a); return; }
            if (avail <= 0) return;
            super.replace(fb, off, len, s.substring(0, Math.min(s.length(), avail)), a);
        }
    }

    private abstract static class DocIns
            implements javax.swing.event.DocumentListener {
        public abstract void onInsert();
        public void insertUpdate(javax.swing.event.DocumentEvent e)  { onInsert(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e)  {}
        public void changedUpdate(javax.swing.event.DocumentEvent e) {}
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
