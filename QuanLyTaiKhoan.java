package quanlynhanvien;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

/**
 * Giao diện quản lý tài khoản
 *
 * @author Windows 10 Version 2
 */
public class QuanLyTaiKhoan extends JFrame {

    // Components
    private JTable table;
    private DefaultTableModel tableModel;
    private JSpinner spinnerNhanVien;
    private JTextField txtMaNhanVien, txtChucVu, txtPhongBan, txtTenDangNhap, txtMatKhau;
    private JButton btnThem, btnSua, btnXoa, btnTimKiem, btnVoHieuHoa, btnBoVoHieuHoa;
    private JTextField txtTimKiem;
    private List<NhanVien> danhSachNhanVien;

    private JComboBox<NhanVien> comboHoTen;
    private JComboBox<String> comboQuyenTruyCap;
    private JComboBox<String> comboTrangThai;
    private JButton btnReset;

    // Inner class for Employee data
    private class NhanVien {

        String maNhanVien;
        String tenNhanVien;
        String phongBan;
        String chucVu;

        public NhanVien(String maNhanVien, String tenNhanVien, String phongBan, String chucVu) {
            this.maNhanVien = maNhanVien;
            this.tenNhanVien = tenNhanVien;
            this.phongBan = phongBan;
            this.chucVu = chucVu;
        }

        @Override
        public String toString() {
            return maNhanVien + " - " + tenNhanVien;
        }
    }

    public QuanLyTaiKhoan() {
        initComponents();
        loadNhanVienData();
        loadTaiKhoanData();
        setupEventListeners();
    }

    private void initComponents() {
        setTitle("Quản lý tài khoản");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1100, 700);
        setLocationRelativeTo(null);

        // Create main panel
        JPanel mainPanel = new JPanel(new BorderLayout());

        // Create table panel
        JPanel tablePanel = createTablePanel();

        // Create form panel
        JPanel formPanel = createFormPanel();

        // Create button panel
        JPanel buttonPanel = createButtonPanel();

        // Add panels to main panel
        mainPanel.add(tablePanel, BorderLayout.CENTER);
        mainPanel.add(formPanel, BorderLayout.WEST);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        add(mainPanel);
    }

    private JPanel createTablePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder("Danh sách tài khoản"));

        // Create table
        String[] columnNames = {"Mã NV", "Họ tên", "Chức vụ", "Phòng ban", "Tên đăng nhập","Mật khẩu", "Quyền truy cập", "Trạng thái"};
        tableModel = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        table = new JTable(tableModel);
        table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        table.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                loadSelectedAccount();
            }
        });

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setPreferredSize(new Dimension(750, 400));

        // Search panel
        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        searchPanel.add(new JLabel("Tìm kiếm:"));
        txtTimKiem = new JTextField(20);
        searchPanel.add(txtTimKiem);

        panel.add(searchPanel, BorderLayout.NORTH);
        panel.add(scrollPane, BorderLayout.CENTER);

        return panel;
    }
private void formWindowClosing(java.awt.event.WindowEvent evt) {                                            
    int choice = JOptionPane.showConfirmDialog(
        this,
        "Bạn có chắc chắn muốn thoát?",
        "Xác nhận thoát",
        JOptionPane.YES_NO_OPTION,
        JOptionPane.QUESTION_MESSAGE
    );
    if (choice == JOptionPane.YES_OPTION) {
       new Home(Session.currentUsername).setVisible(true); // Mở lại JFrame1 (Home)
        this.dispose(); // Đóng cửa sổ hiện tại
    }
}
    private JPanel createFormPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder("Thông tin tài khoản"));
        panel.setPreferredSize(new Dimension(350, 400));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;

        // Khởi tạo danh sách nhân viên
        danhSachNhanVien = new ArrayList<>();

        // ComboBox họ tên nhân viên
        comboHoTen = new JComboBox<>();
        comboHoTen.setPreferredSize(new Dimension(200, 25));
        comboHoTen.addActionListener(e -> {
            NhanVien selected = (NhanVien) comboHoTen.getSelectedItem();
            if (selected != null) {
                txtMaNhanVien.setText(selected.maNhanVien);
                txtChucVu.setText(selected.chucVu);
                txtPhongBan.setText(selected.phongBan);
            }
        });

        // Các field
        txtMaNhanVien = new JTextField(15);
        txtChucVu = new JTextField(15);
        txtPhongBan = new JTextField(15);
        txtTenDangNhap = new JTextField(15);
        txtMatKhau = new JTextField(15);
        comboQuyenTruyCap = new JComboBox<>(new String[]{"User", "Admin"});
        comboTrangThai = new JComboBox<>(new String[]{"Hoạt động", "Vô hiệu hóa"});

        // Không cho sửa
        txtMaNhanVien.setEditable(false);
        txtChucVu.setEditable(false);
        txtPhongBan.setEditable(false);

        // Add vào form
        addFormField(panel, gbc, "Họ tên:", comboHoTen, 0);
        addFormField(panel, gbc, "Mã nhân viên:", txtMaNhanVien, 1);
        addFormField(panel, gbc, "Chức vụ:", txtChucVu, 2);
        addFormField(panel, gbc, "Phòng ban:", txtPhongBan, 3);
        addFormField(panel, gbc, "Tên đăng nhập:", txtTenDangNhap, 4);
        addFormField(panel, gbc, "Mật khẩu:", txtMatKhau, 5);
        addFormField(panel, gbc, "Quyền truy cập:", comboQuyenTruyCap, 6);
        addFormField(panel, gbc, "Trạng thái:", comboTrangThai, 7);

        return panel;
    }

    private void addFormField(JPanel panel, GridBagConstraints gbc, String labelText, JComponent component, int row) {
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel(labelText), gbc);

        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(component, gbc);
        gbc.fill = GridBagConstraints.NONE;
    }

    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout());

        btnThem = new JButton("Thêm");
        btnSua = new JButton("Sửa");
        btnXoa = new JButton("Xóa");
        btnVoHieuHoa = new JButton("Vô hiệu hóa");
        btnBoVoHieuHoa = new JButton("Bỏ vô hiệu hóa");
        btnReset = new JButton("Reset");

        panel.add(btnThem);
        panel.add(btnSua);
        panel.add(btnXoa);
        panel.add(btnVoHieuHoa);
        panel.add(btnBoVoHieuHoa);

        panel.add(btnReset);

        return panel;
    }

    private void setupEventListeners() {
        // Spinner selection listener
       

        // Button listeners
        btnThem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                themTaiKhoan();
            }
        });

        btnSua.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                suaTaiKhoan();
            }
        });

        btnXoa.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                xoaTaiKhoan();
            }
        });

        txtTimKiem.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) {
                timKiemTaiKhoan();
            }

            public void removeUpdate(DocumentEvent e) {
                timKiemTaiKhoan();
            }

            public void changedUpdate(DocumentEvent e) {
                timKiemTaiKhoan();
            }
        });

        btnVoHieuHoa.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                voHieuHoaTaiKhoan();
            }
        });

        btnBoVoHieuHoa.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                boVoHieuHoaTaiKhoan();
            }
        });
        btnReset.addActionListener(e -> clearForm());

    }

    private void loadNhanVienData() {
        try (Connection conn = connection.getConnection()) {
            String sql = "SELECT MaNhanVien, TenNhanVien, PhongBan, ChucVu FROM Tb_NhanVien";
            PreparedStatement stmt = conn.prepareStatement(sql);
            ResultSet rs = stmt.executeQuery();

            danhSachNhanVien.clear();
            comboHoTen.removeAllItems();

            while (rs.next()) {
                NhanVien nv = new NhanVien(
                        rs.getString("MaNhanVien"),
                        rs.getString("TenNhanVien"),
                        rs.getString("PhongBan"),
                        rs.getString("ChucVu")
                );
                danhSachNhanVien.add(nv);
                comboHoTen.addItem(nv); // Thêm vào combobox
            }

            // Nếu có dữ liệu thì set mặc định
            if (!danhSachNhanVien.isEmpty()) {
                comboHoTen.setSelectedIndex(0);
            } else {
                JOptionPane.showMessageDialog(this, "Không tìm thấy dữ liệu nhân viên.");
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi tải dữ liệu nhân viên: " + ex.getMessage());
        }
    }

    private void loadTaiKhoanData() {
        try (Connection conn = connection.getConnection()) {
            String sql = "SELECT * FROM Tb_TaiKhoan";
            PreparedStatement stmt = conn.prepareStatement(sql);
            ResultSet rs = stmt.executeQuery();

            tableModel.setRowCount(0);
            while (rs.next()) {
                Object[] row = {
                    rs.getString("MaNV"),
                    rs.getString("HoTen"),
                    rs.getString("ChucVu"),
                    rs.getString("PhongBan"),
                    rs.getString("TenDangNhap"),
                    rs.getString("MatKhau"),
                    rs.getString("QuyenTruyCap"),
                    rs.getString("TrangThai")
                };
                tableModel.addRow(row);
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi tải dữ liệu tài khoản: " + ex.getMessage());
        }
    }

    private void loadSelectedAccount() {
        int selectedRow = table.getSelectedRow();
        if (selectedRow >= 0) {
            String maNV = (String) tableModel.getValueAt(selectedRow, 0);
            String tenDangNhap = (String) tableModel.getValueAt(selectedRow, 4);
            String quyenTruyCap = (String) tableModel.getValueAt(selectedRow, 5);
            String trangThai = (String) tableModel.getValueAt(selectedRow, 6);

            // Tìm nhân viên tương ứng trong comboHoTen và set selected
            for (int i = 0; i < comboHoTen.getItemCount(); i++) {
                NhanVien nv = comboHoTen.getItemAt(i);
                if (nv.maNhanVien.equals(maNV)) {
                    comboHoTen.setSelectedIndex(i);
                    break;
                }
            }

            txtTenDangNhap.setText(tenDangNhap);
            comboQuyenTruyCap.setSelectedItem(quyenTruyCap);
            comboTrangThai.setSelectedItem(trangThai);

            // Load mật khẩu từ DB
            try (Connection conn = connection.getConnection()) {
                String sql = "SELECT MatKhau FROM Tb_TaiKhoan WHERE MaNV = ?";
                PreparedStatement stmt = conn.prepareStatement(sql);
                stmt.setString(1, maNV);
                ResultSet rs = stmt.executeQuery();

                if (rs.next()) {
                    txtMatKhau.setText(rs.getString("MatKhau"));
                }

            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(this, "Lỗi khi tải mật khẩu: " + ex.getMessage());
            }
            table.scrollRectToVisible(table.getCellRect(selectedRow, 0, true));

        }
    }

    private void themTaiKhoan() {
        if (!validateInput()) {
            return;
        }

        NhanVien selectedNV = (NhanVien) comboHoTen.getSelectedItem();
        if (selectedNV == null) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn nhân viên!");
            return;
        }

        try (Connection conn = connection.getConnection()) {
            // Kiểm tra trùng tên đăng nhập hoặc nhân viên đã có tài khoản
            String checkSql = "SELECT COUNT(*) FROM Tb_TaiKhoan WHERE MaNV = ? OR TenDangNhap = ?";
            PreparedStatement checkStmt = conn.prepareStatement(checkSql);
            checkStmt.setString(1, selectedNV.maNhanVien);
            checkStmt.setString(2, txtTenDangNhap.getText().trim());
            ResultSet rs = checkStmt.executeQuery();

            if (rs.next() && rs.getInt(1) > 0) {
                JOptionPane.showMessageDialog(this, "Tên đăng nhập đã tồn tại hoặc nhân viên đã có tài khoản!");
                return;
            }

            // Thêm tài khoản mới
            String sql = "INSERT INTO Tb_TaiKhoan (MaNV, HoTen, ChucVu, PhongBan, TenDangNhap, MatKhau, QuyenTruyCap, TrangThai) VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
            PreparedStatement stmt = conn.prepareStatement(sql);

            stmt.setString(1, selectedNV.maNhanVien);
            stmt.setString(2, selectedNV.tenNhanVien);
            stmt.setString(3, selectedNV.chucVu);
            stmt.setString(4, selectedNV.phongBan);
            stmt.setString(5, txtTenDangNhap.getText().trim());
            stmt.setString(6, txtMatKhau.getText().trim());
            stmt.setString(7, comboQuyenTruyCap.getSelectedItem().toString());
            stmt.setString(8, comboTrangThai.getSelectedItem().toString());

            int result = stmt.executeUpdate();
            if (result > 0) {
                JOptionPane.showMessageDialog(this, "Thêm tài khoản thành công!");
                loadTaiKhoanData();
                clearForm();
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi thêm tài khoản: " + ex.getMessage());
        }
    }

    private void suaTaiKhoan() {
        int selectedRow = table.getSelectedRow();
        if (selectedRow < 0) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn tài khoản cần sửa!");
            return;
        }

        if (!validateInput()) {
            return;
        }

        NhanVien selectedNV = (NhanVien) comboHoTen.getSelectedItem();
        if (selectedNV == null) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn nhân viên!");
            return;
        }

        String originalMaNV = (String) tableModel.getValueAt(selectedRow, 0);

        try (Connection conn = connection.getConnection()) {
            // Kiểm tra trùng tên đăng nhập (trừ tài khoản hiện tại)
            String checkSql = "SELECT COUNT(*) FROM Tb_TaiKhoan WHERE TenDangNhap = ? AND MaNV != ?";
            PreparedStatement checkStmt = conn.prepareStatement(checkSql);
            checkStmt.setString(1, txtTenDangNhap.getText().trim());
            checkStmt.setString(2, originalMaNV);
            ResultSet rs = checkStmt.executeQuery();

            if (rs.next() && rs.getInt(1) > 0) {
                JOptionPane.showMessageDialog(this, "Tên đăng nhập đã được sử dụng!");
                return;
            }

            // Cập nhật dữ liệu tài khoản
            String sql = "UPDATE Tb_TaiKhoan SET HoTen = ?, ChucVu = ?, PhongBan = ?, TenDangNhap = ?, MatKhau = ?, QuyenTruyCap = ?, TrangThai = ? WHERE MaNV = ?";
            PreparedStatement stmt = conn.prepareStatement(sql);

            stmt.setString(1, selectedNV.tenNhanVien);
            stmt.setString(2, selectedNV.chucVu);
            stmt.setString(3, selectedNV.phongBan);
            stmt.setString(4, txtTenDangNhap.getText().trim());
            stmt.setString(5, txtMatKhau.getText().trim());
            stmt.setString(6, comboQuyenTruyCap.getSelectedItem().toString());
            stmt.setString(7, comboTrangThai.getSelectedItem().toString());
            stmt.setString(8, originalMaNV);

            int result = stmt.executeUpdate();
            if (result > 0) {
                JOptionPane.showMessageDialog(this, "Cập nhật tài khoản thành công!");
                loadTaiKhoanData();
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi cập nhật tài khoản: " + ex.getMessage());
        }
    }

    private void xoaTaiKhoan() {
        int selectedRow = table.getSelectedRow();
        if (selectedRow < 0) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn tài khoản cần xóa!");
            return;
        }

        int confirm = JOptionPane.showConfirmDialog(this,
                "Bạn có chắc chắn muốn xóa tài khoản này?",
                "Xác nhận xóa",
                JOptionPane.YES_NO_OPTION);

        if (confirm == JOptionPane.YES_OPTION) {
            String maNV = (String) tableModel.getValueAt(selectedRow, 0);

            try (Connection conn = connection.getConnection()) {
                String sql = "DELETE FROM Tb_TaiKhoan WHERE MaNV = ?";
                PreparedStatement stmt = conn.prepareStatement(sql);
                stmt.setString(1, maNV);

                int result = stmt.executeUpdate();
                if (result > 0) {
                    JOptionPane.showMessageDialog(this, "Xóa tài khoản thành công!");
                    loadTaiKhoanData();
                    clearForm();
                }

            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(this, "Lỗi khi xóa tài khoản: " + ex.getMessage());
            }
        }
    }

    private void clearForm() {
        if (!danhSachNhanVien.isEmpty()) {
            comboHoTen.setSelectedIndex(0);
        }
        txtTenDangNhap.setText("");
        txtMatKhau.setText("");
        comboQuyenTruyCap.setSelectedIndex(0);
        comboTrangThai.setSelectedIndex(0);
        txtTenDangNhap.requestFocus();

    }

    private void timKiemTaiKhoan() {
        String keyword = txtTimKiem.getText().trim();
        if (keyword.isEmpty()) {
            loadTaiKhoanData();
            return;
        }

        try (Connection conn = connection.getConnection()) {
            String sql = "SELECT * FROM Tb_TaiKhoan WHERE MaNV LIKE ? OR HoTen LIKE ? OR ChucVu LIKE ? OR PhongBan LIKE ? OR TenDangNhap LIKE ? OR QuyenTruyCap LIKE ? OR TrangThai LIKE ?";
            PreparedStatement stmt = conn.prepareStatement(sql);

            String searchPattern = "%" + keyword + "%";
            for (int i = 1; i <= 7; i++) {
                stmt.setString(i, searchPattern);
            }

            ResultSet rs = stmt.executeQuery();

            tableModel.setRowCount(0);
            while (rs.next()) {
                Object[] row = {
                    rs.getString("MaNV"),
                    rs.getString("HoTen"),
                    rs.getString("ChucVu"),
                    rs.getString("PhongBan"),
                    rs.getString("TenDangNhap"),
                    rs.getString("QuyenTruyCap"),
                    rs.getString("TrangThai")
                };
                tableModel.addRow(row);
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi tìm kiếm: " + ex.getMessage());
        }
    }

    private void voHieuHoaTaiKhoan() {
        updateTrangThaiTaiKhoan("Vô hiệu hóa");
    }

    private void boVoHieuHoaTaiKhoan() {
        updateTrangThaiTaiKhoan("Hoạt động");
    }

private void updateTrangThaiTaiKhoan(String trangThai) {
    int selectedRow = table.getSelectedRow();
    if (selectedRow < 0) {
        JOptionPane.showMessageDialog(this, "Vui lòng chọn tài khoản cần thay đổi trạng thái!");
        return;
    }

    String maNV = (String) tableModel.getValueAt(selectedRow, 0);
    String lyDo = JOptionPane.showInputDialog(this, "Nhập lý do thay đổi trạng thái:");

    if (lyDo != null && !lyDo.trim().isEmpty()) {
        try (Connection conn = connection.getConnection()) {
            String sql = "UPDATE Tb_TaiKhoan SET TrangThai = ? WHERE MaNV = ?";
            PreparedStatement stmt = conn.prepareStatement(sql);
            stmt.setString(1, trangThai);
            stmt.setString(2, maNV);

            int result = stmt.executeUpdate();
            if (result > 0) {
                JOptionPane.showMessageDialog(this, "Cập nhật trạng thái thành công!");
                loadTaiKhoanData();
                comboTrangThai.setSelectedItem(trangThai); // cập nhật comboBox
            }

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Lỗi khi cập nhật trạng thái: " + ex.getMessage());
        }
    }
}


  private boolean validateInput() {
    if (comboHoTen.getSelectedItem() == null) {
        JOptionPane.showMessageDialog(this, "Vui lòng chọn họ tên nhân viên!");
        return false;
    }

    if (txtTenDangNhap.getText().trim().isEmpty()) {
        JOptionPane.showMessageDialog(this, "Tên đăng nhập không được để trống!");
        return false;
    }

    if (txtMatKhau.getText().trim().isEmpty()) {
        JOptionPane.showMessageDialog(this, "Mật khẩu không được để trống!");
        return false;
    }

    if (txtMatKhau.getText().trim().length() < 6) {
        JOptionPane.showMessageDialog(this, "Mật khẩu phải có ít nhất 6 ký tự!");
        return false;
    }

    if (comboQuyenTruyCap.getSelectedItem() == null || comboQuyenTruyCap.getSelectedItem().toString().trim().isEmpty()) {
        JOptionPane.showMessageDialog(this, "Vui lòng chọn quyền truy cập!");
        return false;
    }

    if (comboTrangThai.getSelectedItem() == null || comboTrangThai.getSelectedItem().toString().trim().isEmpty()) {
        JOptionPane.showMessageDialog(this, "Vui lòng chọn trạng thái!");
        return false;
    }

    return true;
}


    

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new QuanLyTaiKhoan().setVisible(true);
        });
    }
}
