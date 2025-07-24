/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JPanel.java to edit this template
 */
package com.raven.form;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import javax.swing.BorderFactory;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
public class JfKhenThuong extends javax.swing.JPanel {

    private Map<String, Integer> soTienKT = new HashMap<>();

    /**
     * Creates new form JfKhenThuong
     */
    public JfKhenThuong() {
        initComponents();
        initComboBoxData();

        cboLyDoKT.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent evt) {
                cboLyDoKTItemStateChanged(evt);
            }
        });

        txtTimkiem.setText("Tìm kiếm ở đây...");
        txtTimkiem.setForeground(Color.GRAY);

        txtTimkiem.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                if (txtTimkiem.getText().equals("Tìm kiếm ở đây...")) {
                    txtTimkiem.setText("");
                    txtTimkiem.setForeground(Color.BLACK);
                }
            }

            public void focusLost(java.awt.event.FocusEvent evt) {
                if (txtTimkiem.getText().isEmpty()) {
                    txtTimkiem.setText("Tìm kiếm ở đây...");
                    txtTimkiem.setForeground(Color.GRAY);
                }
            }
        });
        txtTimkiem.addKeyListener(new KeyAdapter() {
            @Override
            public void keyReleased(KeyEvent e) {
                timkiemkhenthuong(txtTimkiem.getText().trim());
            }
        });

        cbotennv.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String tenNV = (String) cbotennv.getSelectedItem();
                if (tenNV != null && !tenNV.equals("-- Chọn nhân viên --")) {
                    loadThongTinNhanVien();
                } else {
                    clearThongTinNhanVien();
                }
            }
        });
        loadTenNhanVien();
        laydulieu();
        jTable1.getSelectionModel().addListSelectionListener(e -> {
            // Đảm bảo không gọi 2 lần khi thay đổi cột
            if (!e.getValueIsAdjusting()) {
                Nhapdulieu();
            }
        });
    }

    public void clearForm() {
        txtMakt.setText("");
        cbotennv.setSelectedIndex(0);
        lbmanv.setText("");
        lbpb.setText("");
        lbchucvu.setText("");
        DateNgaykth.setDate(null);
        cboLyDoKT.setSelectedIndex(0);
        lblkt.setText("");
    }

    private void clearThongTinNhanVien() {
        lbmanv.setText("");
        lbpb.setText("");
        lbchucvu.setText("");
        lblkt.setText("");
    }

    private void Nhapdulieu() {
        int selectedRow = jTable1.getSelectedRow();

        if (selectedRow != -1) {

            String maKhenThuong = jTable1.getValueAt(selectedRow, 0).toString();
            String tenNhanVien = jTable1.getValueAt(selectedRow, 2).toString();
            String maNhanVien = jTable1.getValueAt(selectedRow, 1).toString();
            String PhongBan = jTable1.getValueAt(selectedRow, 3).toString();
            String ChucVu = jTable1.getValueAt(selectedRow, 4).toString();
            Date ngayKt = (Date) jTable1.getValueAt(selectedRow, 5);
            String khenThuong = jTable1.getValueAt(selectedRow, 6).toString();
            String lyDokt = jTable1.getValueAt(selectedRow, 7).toString();

            txtMakt.setText(maKhenThuong);
            cbotennv.setSelectedItem(tenNhanVien);
            lbmanv.setText(maNhanVien);
            lbpb.setText(PhongBan);
            lbchucvu.setText(ChucVu);
            DateNgaykth.setDate(ngayKt);

            lblkt.setText(khenThuong);
            cboLyDoKT.setSelectedItem(lyDokt);

        }
    }

    private void laydulieu() {
        String sql = "SELECT MaKT, MaNhanVien, TenNhanVien, PhongBan, ChucVu, NgayKT, KhenThuong, LyDoKT FROM Tb_KhenThuong";

        try (Connection conn = connection.getConnection(); PreparedStatement ps = conn.prepareStatement(sql)) {
            ResultSet rs = ps.executeQuery();

            // Đặt tiêu đề cột đầy đủ
            DefaultTableModel tableModel = new DefaultTableModel(
                    new Object[]{"Mã KT", "Mã NV", "Tên nhân viên", "Phòng ban", "Chức vụ", "Ngày KT", "Khen thưởng", "Lý do KT"}, 0
            );
            jTable1.setModel(tableModel);

            while (rs.next()) {
                Object[] row = new Object[8];
                row[0] = rs.getString("MaKT");
                row[1] = rs.getString("MaNhanVien");
                row[2] = rs.getString("TenNhanVien");
                row[3] = rs.getString("PhongBan");
                row[4] = rs.getString("ChucVu");
                row[5] = rs.getDate("NgayKT");
                row[6] = rs.getString("KhenThuong");
                row[7] = rs.getString("LyDoKT");

                tableModel.addRow(row);
            }

        } catch (SQLException e) {
            e.printStackTrace();
            ThongBao("Lỗi khi tải dữ liệu khen thưởng.", "Lỗi", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void initComboBoxData() {

        soTienKT.put("Hoàn thành xuất sắc", 1000000);
        soTienKT.put("Dẫn đầu doanh số", 1500000);
        soTienKT.put("Chuyên cần", 500000);
        soTienKT.put("Sáng kiến cải tiến", 1200000);
        soTienKT.put("Hỗ trợ đồng nghiệp tốt", 700000);
        soTienKT.put("Thái độ làm việc tích cực", 800000);
        soTienKT.put("Làm việc ngoài giờ", 600000);
        soTienKT.put("Tăng ca lễ tết", 1000000);
        soTienKT.put("Hoàn thành dự án đúng hạn", 1300000);
        soTienKT.put("Đạt chứng chỉ chuyên môn", 1100000);
        soTienKT.put("Khác", 0); // Cho phép người dùng nhập

        // Thêm từng lý do vào combo
        for (String lyDo : soTienKT.keySet()) {
            cboLyDoKT.addItem(lyDo);
        }
    }

    private void cboLyDoKTItemStateChanged(java.awt.event.ItemEvent evt) {
        if (evt.getStateChange() == ItemEvent.SELECTED) {
            String lyDo = cboLyDoKT.getSelectedItem().toString();

            if (lyDo.equals("Khác")) {
                String nhap = JOptionPane.showInputDialog(this, "Nhập số tiền khen thưởng:");
                if (nhap != null && !nhap.trim().isEmpty()) {
                    try {
                        String cleaned = nhap.replaceAll("[^\\d]", "");
                        int soTien = Integer.parseInt(cleaned);

                        lblkt.setText(String.format("%,d VNĐ", soTien));
                        lblkt.putClientProperty("rawValue", soTien);
                    } catch (NumberFormatException e) {
                        JOptionPane.showMessageDialog(this, "Số tiền không hợp lệ.", "Lỗi", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    lblkt.setText("");
                }
            } else {
                int soTien = soTienKT.getOrDefault(lyDo, 0);
                lblkt.setText(String.format("%,d VNĐ", soTien));
                lblkt.putClientProperty("rawValue", soTien);
            }
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        txtMakt = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        cbotennv = new javax.swing.JComboBox<>();
        lbmanv = new javax.swing.JLabel();
        cboLyDoKT = new javax.swing.JComboBox<>();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        lbpb = new javax.swing.JLabel();
        lbchucvu = new javax.swing.JLabel();
        lblkt = new javax.swing.JLabel();
        btnThem = new javax.swing.JButton();
        btnSua = new javax.swing.JButton();
        btnXoa = new javax.swing.JButton();
        btnReset = new javax.swing.JButton();
        btnXuat = new javax.swing.JButton();
        jLabel8 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        txtTimkiem = new javax.swing.JTextField();

        setPreferredSize(new java.awt.Dimension(1197, 762));

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));
        jPanel1.setPreferredSize(new java.awt.Dimension(1197, 762));

        jPanel2.setBackground(new java.awt.Color(255, 255, 255));
        jPanel2.setPreferredSize(new java.awt.Dimension(379, 762));

        txtMakt.setFont(new java.awt.Font("SansSerif", 0, 12)); // NOI18N
        txtMakt.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        jLabel1.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel1.setText("Mã khen thưởng");

        jLabel2.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel2.setText("Mã nhân viên");

        jLabel3.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel3.setText("Tên nhân viên");

        jLabel4.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel4.setText("Khen thưởng");

        jLabel5.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel5.setText("Ngày khen thưởng");

        jLabel6.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel6.setText("Lý do khen thưởng");

        cbotennv.setBorder(new javax.swing.border.MatteBorder(null));

        lbmanv.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbmanv.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lbmanv.setText("Mã nhân viên");
        lbmanv.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);

        cboLyDoKT.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "None" }));
        cboLyDoKT.setBorder(new javax.swing.border.MatteBorder(null));
        cboLyDoKT.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cboLyDoKTActionPerformed(evt);
            }
        });

        jLabel9.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel9.setText("Phòng ban");

        jLabel10.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel10.setText("Chức vụ");

        lbpb.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbpb.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lbpb.setText("Phòng ban");
        lbpb.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);

        lbchucvu.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbchucvu.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lbchucvu.setText("Chức vụ");
        lbchucvu.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);

        lblkt.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lblkt.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblkt.setText("Khen thưởng");
        lblkt.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);

        btnThem.setBackground(new java.awt.Color(66, 139, 202));
        btnThem.setFont(new java.awt.Font("SansSerif", 1, 12)); // NOI18N
        btnThem.setText("Thêm");
        btnThem.setPreferredSize(new java.awt.Dimension(85, 40));
        btnThem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnThemActionPerformed(evt);
            }
        });

        btnSua.setBackground(new java.awt.Color(66, 139, 202));
        btnSua.setFont(new java.awt.Font("SansSerif", 1, 12)); // NOI18N
        btnSua.setText("Sửa");
        btnSua.setPreferredSize(new java.awt.Dimension(85, 40));
        btnSua.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSuaActionPerformed(evt);
            }
        });

        btnXoa.setBackground(new java.awt.Color(66, 139, 202));
        btnXoa.setFont(new java.awt.Font("SansSerif", 1, 12)); // NOI18N
        btnXoa.setText("Xoá");
        btnXoa.setPreferredSize(new java.awt.Dimension(85, 40));
        btnXoa.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnXoaActionPerformed(evt);
            }
        });

        btnReset.setBackground(new java.awt.Color(66, 139, 202));
        btnReset.setFont(new java.awt.Font("SansSerif", 1, 12)); // NOI18N
        btnReset.setText("Làm mới");
        btnReset.setPreferredSize(new java.awt.Dimension(85, 40));
        btnReset.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnResetActionPerformed(evt);
            }
        });

        btnXuat.setBackground(new java.awt.Color(66, 139, 202));
        btnXuat.setFont(new java.awt.Font("SansSerif", 1, 12)); // NOI18N
        btnXuat.setText("Xuất File");
        btnXuat.setPreferredSize(new java.awt.Dimension(85, 40));
        btnXuat.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnXuatActionPerformed(evt);
            }
        });

        jLabel8.setFont(new java.awt.Font("Arial", 0, 24)); // NOI18N
        jLabel8.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/raven/icon/gold-medal.png"))); // NOI18N
        jLabel8.setText("Khen Thưởng");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel3)
                            .addComponent(jLabel10)
                            .addComponent(jLabel2)
                            .addComponent(jLabel5)
                            .addComponent(jLabel4)
                            .addComponent(jLabel6))
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(lblkt, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                        .addComponent(lbchucvu, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(lbpb, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(lbmanv, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(cbotennv, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(txtMakt, javax.swing.GroupLayout.Alignment.LEADING))
                                    .addComponent(cboLyDoKT, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGap(47, 47, 47))
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGap(97, 97, 97)
                                .addComponent(btnXoa, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel1)
                            .addComponent(jLabel9)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addComponent(btnThem, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(btnSua, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(48, 48, 48)
                        .addComponent(btnReset, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(54, 54, 54)
                        .addComponent(btnXuat, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(65, 65, 65)
                        .addComponent(jLabel8)))
                .addGap(0, 0, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addGap(38, 38, 38)
                .addComponent(jLabel8)
                .addGap(67, 67, 67)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(txtMakt, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(32, 32, 32)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(cbotennv, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(35, 35, 35)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(lbmanv))
                .addGap(32, 32, 32)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel9)
                    .addComponent(lbpb))
                .addGap(37, 37, 37)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel10)
                    .addComponent(lbchucvu))
                .addGap(35, 35, 35)
                .addComponent(jLabel5)
                .addGap(44, 44, 44)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(lblkt))
                .addGap(47, 47, 47)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(cboLyDoKT, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(36, 36, 36)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnThem, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnSua, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnXoa, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(27, 27, 27)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnReset, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnXuat, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(78, 78, 78))
        );

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        Font textFont = new Font("SansSerif", Font.PLAIN, 12);

        Color textBorder = new Color(200, 200, 200);

        txtTimkiem.setFont(textFont);
        txtTimkiem.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(textBorder),
            BorderFactory.createEmptyBorder(5, 5, 5, 5)
        ));

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(359, 359, 359)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 838, Short.MAX_VALUE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(txtTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, 400, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))))
            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel1Layout.createSequentialGroup()
                    .addContainerGap()
                    .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, 350, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(841, Short.MAX_VALUE)))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(txtTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1)
                .addContainerGap())
            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel1Layout.createSequentialGroup()
                    .addContainerGap()
                    .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
    }// </editor-fold>                        

    private void btnThemActionPerformed(java.awt.event.ActionEvent evt) {                                        
        themKhenThuong();
    }                                       

    private void btnSuaActionPerformed(java.awt.event.ActionEvent evt) {                                       
        suaKhenThuong();
    }                                      

    private void btnXoaActionPerformed(java.awt.event.ActionEvent evt) {                                       

        int selectedRow = jTable1.getSelectedRow();

        if (selectedRow == -1) {
            ThongBao("Vui lòng chọn một khen thưởng để xóa!", "Lỗi", JOptionPane.ERROR_MESSAGE);
            return;
        }

        String makt = jTable1.getValueAt(selectedRow, 0).toString();

        int confirm = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn muốn xóa khen thưởng này?", "Xác nhận xóa", JOptionPane.YES_NO_OPTION);

        if (confirm == JOptionPane.YES_OPTION) {

            try (Connection conn = connection.getConnection()) {
                String deleteSql = "DELETE FROM Tb_KhenThuong WHERE MaKT = ?";
                PreparedStatement psDelete = conn.prepareStatement(deleteSql);
                psDelete.setString(1, makt);

                int rowsAffected = psDelete.executeUpdate();

                if (rowsAffected > 0) {
                    ThongBao("Đã xóa khen thưởng thành công!", "Thông báo", JOptionPane.INFORMATION_MESSAGE);
                    laydulieu();
                } else {
                    ThongBao("Không tìm thấy khen thưởng với mã đã chọn!", "Thông báo", JOptionPane.WARNING_MESSAGE);
                }
            } catch (Exception e) {
                System.out.println(e.toString());
                ThongBao("Có lỗi xảy ra trong quá trình xóa khen thưởng.", "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        } else {

            ThongBao("Hủy thao tác xóa!", "Thông báo", JOptionPane.INFORMATION_MESSAGE);
        }
    }                                      

    private void btnResetActionPerformed(java.awt.event.ActionEvent evt) {                                         
        clearForm();
    }                                        

    private void btnXuatActionPerformed(java.awt.event.ActionEvent evt) {                                        
        String[] options = {"Xuất PDF", "Xuất Excel", "Hủy"};
        int choice = JOptionPane.showOptionDialog(this,
                "Chọn định dạng bạn muốn xuất:",
                "Xuất dữ liệu",
                JOptionPane.DEFAULT_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                options,
                options[0]);

        if (choice == 0) {
            xuatPDF();
        } else if (choice == 1) {
            xuatExcel();
        }

    }                                       

    private void cboLyDoKTActionPerformed(java.awt.event.ActionEvent evt) {                                          
        // TODO add your handling code here:
    }                                         
    private void xuatPDF() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu file PDF");
        fileChooser.setSelectedFile(new File("Khenthuong.pdf"));
        int userSelection = fileChooser.showSaveDialog(this);

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            try {
                Document document = new Document();
                PdfWriter.getInstance(document, new FileOutputStream(fileToSave));
                document.open();

                PdfPTable pdfTable = new PdfPTable(jTable1.getColumnCount());
                // Thêm tiêu đề
                for (int i = 0; i < jTable1.getColumnCount(); i++) {
                    pdfTable.addCell(new PdfPCell(new Phrase(jTable1.getColumnName(i))));
                }
                // Thêm dữ liệu
                for (int row = 0; row < jTable1.getRowCount(); row++) {
                    for (int col = 0; col < jTable1.getColumnCount(); col++) {
                        Object value = jTable1.getValueAt(row, col);
                        pdfTable.addCell(value != null ? value.toString() : "");
                    }
                }

                document.add(pdfTable);
                document.close();
                ThongBao("Xuất PDF thành công!", "Thông báo", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ex.printStackTrace();
                ThongBao("Lỗi khi xuất PDF: " + ex.getMessage(), "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void xuatExcel() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu file Excel");
        fileChooser.setSelectedFile(new File("khenthuong.xlsx"));
        int userSelection = fileChooser.showSaveDialog(this);

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("DanhSachKhenThuong");

                // Tạo hàng tiêu đề
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < jTable1.getColumnCount(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(jTable1.getColumnName(i));
                }

                // Thêm dữ liệu
                for (int row = 0; row < jTable1.getRowCount(); row++) {
                    Row excelRow = sheet.createRow(row + 1);
                    for (int col = 0; col < jTable1.getColumnCount(); col++) {
                        Object value = jTable1.getValueAt(row, col);
                        Cell cell = excelRow.createCell(col);
                        cell.setCellValue(value != null ? value.toString() : "");
                    }
                }

                FileOutputStream fileOut = new FileOutputStream(fileToSave);
                workbook.write(fileOut);
                fileOut.close();
                ThongBao("Xuất Excel thành công!", "Thông báo", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                ex.printStackTrace();
                ThongBao("Lỗi khi xuất Excel: " + ex.getMessage(), "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void ThongBao(String noiDungThongBao, String tieuDeThongBao, int icon) {
        JOptionPane.showMessageDialog(new JFrame(), noiDungThongBao,
                tieuDeThongBao, icon);
    }

    private void loadTenNhanVien() {
        try {
            cbotennv.removeAllItems();
            cbotennv.addItem("-- Chọn nhân viên --");

            String sql = "SELECT MaNhanVien, TenNhanVien FROM Tb_NhanVien ORDER BY TenNhanVien";
            Connection conn = connection.getConnection();
            PreparedStatement pst = conn.prepareStatement(sql);
            ResultSet rs = pst.executeQuery();

            while (rs.next()) {
                cbotennv.addItem(rs.getString("MaNhanVien") + " - " + rs.getString("TenNhanVien"));
            }

            rs.close();
            pst.close();
            conn.close();
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Lỗi load danh sách nhân viên: " + e.getMessage());
        }
    }

    private void loadThongTinNhanVien() {
        String selectedItem = cbotennv.getSelectedItem().toString();

        // Giả sử format là "NV01 - Nguyễn Văn A"
        if (!selectedItem.contains(" - ")) {
            return; // Tránh lỗi nếu chọn "-- Chọn nhân viên --"
        }
        String maNhanVien = selectedItem.split(" - ")[0];

        String sql = "SELECT MaNhanVien, PhongBan, ChucVu FROM Tb_NhanVien WHERE MaNhanVien = ?";

        try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {

            pst.setString(1, maNhanVien);
            try (ResultSet rs = pst.executeQuery()) {
                if (rs.next()) {
                    lbmanv.setText(rs.getString("MaNhanVien"));
                    lbpb.setText(rs.getString("PhongBan")); // Hiển thị mã phòng ban (nếu cần tên, bạn phải JOIN)
                    lbchucvu.setText(rs.getString("ChucVu"));
                }
            }

        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Lỗi load thông tin nhân viên: " + e.getMessage());
        }
    }

    private boolean kiemTraMaKTTonTai(String maKT) {
        String sql = "SELECT COUNT(*) FROM Tb_KhenThuong WHERE MaKT = ?";

        try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {

            pst.setString(1, maKT);
            try (ResultSet rs = pst.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt(1) > 0;
                }
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return false;
    }

    private boolean validateInput() {
        if (txtMakt.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "Vui lòng nhập mã khen thưởng!");
            txtMakt.requestFocus();
            return false;
        }

        if (cbotennv.getSelectedIndex() == 0) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn nhân viên!");
            cbotennv.requestFocus();
            return false;
        }

        if (DateNgaykth.getDate() == null) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn ngày khen thưởng!");
            DateNgaykth.requestFocus();
            return false;
        }

        String lyDo = (String) cboLyDoKT.getSelectedItem();
        if (lyDo == null || lyDo.equalsIgnoreCase("None") || lyDo.equals("-- Chọn lý do --") || cboLyDoKT.getSelectedIndex() == 0) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn lý do khen thưởng!");
            cboLyDoKT.requestFocus();
            return false;
        }

        if (lblkt.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "Vui lòng nhập lý do khen thưởng!");
            lblkt.requestFocus();
            return false;
        }

        return true;
    }

    private void timkiemkhenthuong(String keyword) {
        if (keyword == null || keyword.trim().isEmpty()) {
            laydulieu(); // Gọi lại hàm hiển thị toàn bộ nếu không có từ khóa
            return;
        }

        String sql = "SELECT * FROM Tb_KhenThuong WHERE MaKT LIKE ? OR MaNhanVien LIKE ? OR TenNhanVien LIKE ? OR KhenThuong LIKE ?";

        try (Connection conn = connection.getConnection(); PreparedStatement ps = conn.prepareStatement(sql)) {
            String searchPattern = "%" + keyword + "%";
            for (int i = 1; i <= 4; i++) {
                ps.setString(i, searchPattern);
            }

            DefaultTableModel tableModel = new DefaultTableModel(
                    new Object[]{"Mã KT", "Mã NV", "Tên nhân viên", "Phòng ban", "Chức vụ", "Ngày KT", "Khen thưởng", "Lý do KT"}, 0
            );
            jTable1.setModel(tableModel);

            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                Object[] item = new Object[8];
                item[0] = rs.getString("MaKT");
                item[1] = rs.getString("MaNhanVien");
                item[2] = rs.getString("TenNhanVien");
                item[3] = rs.getString("PhongBan"); // Vẫn hiển thị mã phòng ban
                item[4] = rs.getString("ChucVu");
                item[5] = rs.getDate("NgayKT");
                item[6] = rs.getString("KhenThuong");
                item[7] = rs.getString("LyDoKT");

                tableModel.addRow(item);
            }
        } catch (SQLException e) {
            e.printStackTrace();
            ThongBao("Có lỗi khi tìm kiếm dữ liệu khen thưởng.", "Lỗi", JOptionPane.ERROR_MESSAGE);
        }
    }

private void themKhenThuong() {
    if (!validateInput()) {
        return;
    }
    if (kiemTraMaKTTonTai(txtMakt.getText().trim())) {
        JOptionPane.showMessageDialog(null, "Mã khen thưởng đã tồn tại!");
        return;
    }
    
    String sql = "INSERT INTO Tb_KhenThuong (MaKT, MaNhanVien, TenNhanVien, PhongBan, ChucVu, NgayKT, KhenThuong, LyDoKT) VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
    
    try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {
        pst.setString(1, txtMakt.getText().trim());                     // MaKT
        pst.setString(2, lbmanv.getText().trim());                     // MaNhanVien
        
        // Tách tên nhân viên từ combo box: "NV01 - Nguyễn Văn A" -> "Nguyễn Văn A"
        String selected = (String) cbotennv.getSelectedItem();
        String tenNhanVien = selected.contains(" - ") ? selected.split(" - ", 2)[1] : selected;
        pst.setString(3, tenNhanVien);                                 // TenNhanVien
        
        pst.setString(4, lbpb.getText().trim());                       // PhongBan
        pst.setString(5, lbchucvu.getText().trim());                   // ChucVu
        
        // Ngày khen thưởng
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String ngayKT = sdf.format(DateNgaykth.getDate());
        pst.setString(6, ngayKT);
        
        // Lấy lý do khen thưởng
        String lyDoKT = (String) cboLyDoKT.getSelectedItem();
        
        // Lấy số tiền khen thưởng
        int soTienKT;
        if (lyDoKT.equals("Khác")) {
            // Lấy giá trị raw từ clientProperty hoặc parse từ label
            Object rawValue = lblkt.getClientProperty("rawValue");
            if (rawValue != null && rawValue instanceof Integer) {
                soTienKT = (Integer) rawValue;
            } else {
                // Fallback: parse từ text của label
                String soTienText = lblkt.getText().trim();
                if (soTienText.isEmpty() || soTienText.equals("")) {
                    JOptionPane.showMessageDialog(null, "Vui lòng nhập số tiền khen thưởng!");
                    return;
                }
                try {
                    // Loại bỏ định dạng tiền tệ và parse
                    String cleaned = soTienText.replaceAll("[^\\d]", "");
                    soTienKT = Integer.parseInt(cleaned);
                } catch (NumberFormatException ex) {
                    JOptionPane.showMessageDialog(null, "Số tiền không hợp lệ!");
                    return;
                }
            }
        } else {
            // Lấy từ HashMap soTienKT (HashMap bạn khai báo ở đầu class)
            soTienKT = this.soTienKT.getOrDefault(lyDoKT, 0);
        }
        
        pst.setInt(7, soTienKT);                                       // Số tiền khen thưởng
        pst.setString(8, lyDoKT);                                      // Lý do khen thưởng
        
        int result = pst.executeUpdate();
        if (result > 0) {
            JOptionPane.showMessageDialog(null, "Thêm khen thưởng thành công!");
            clearForm();
        } else {
            JOptionPane.showMessageDialog(null, "Thêm khen thưởng thất bại!");
        }
        
    } catch (SQLException e) {
        e.printStackTrace();
        JOptionPane.showMessageDialog(null, "Lỗi thêm khen thưởng: " + e.getMessage());
    }
    
    laydulieu();
}

    private void suaKhenThuong() {
    if (!validateInput()) {
        return;
    }
    
    // Kiểm tra xem có dòng nào được chọn không
    int selectedRow = jTable1.getSelectedRow();
    if (selectedRow == -1) {
        JOptionPane.showMessageDialog(null, "Vui lòng chọn một dòng để sửa!");
        return;
    }
    
    String sql = "UPDATE Tb_KhenThuong SET MaNhanVien=?, TenNhanVien=?, PhongBan=?, ChucVu=?, NgayKT=?, KhenThuong=?, LyDoKT=? WHERE MaKT=?";
    
    try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {
        pst.setString(1, lbmanv.getText().trim());                     // MaNhanVien
        
        // Tách tên nhân viên từ combo box: "NV01 - Nguyễn Văn A" -> "Nguyễn Văn A"
        String selected = (String) cbotennv.getSelectedItem();
        String tenNhanVien = selected.contains(" - ") ? selected.split(" - ", 2)[1] : selected;
        pst.setString(2, tenNhanVien);                                 // TenNhanVien
        
        pst.setString(3, lbpb.getText().trim());                       // PhongBan
        pst.setString(4, lbchucvu.getText().trim());                   // ChucVu
        
        // Ngày khen thưởng
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String ngayKT = sdf.format(DateNgaykth.getDate());
        pst.setString(5, ngayKT);
        
        // Lấy lý do khen thưởng
        String lyDoKT = (String) cboLyDoKT.getSelectedItem();
        
        // Lấy số tiền khen thưởng
        int soTien;
        if (lyDoKT.equals("Khác")) {
            // Lấy giá trị raw từ clientProperty hoặc parse từ label
            Object rawValue = lblkt.getClientProperty("rawValue");
            if (rawValue != null && rawValue instanceof Integer) {
                soTien = (Integer) rawValue;
            } else {
                // Fallback: parse từ text của label
                String soTienText = lblkt.getText().trim();
                if (soTienText.isEmpty() || soTienText.equals("")) {
                    JOptionPane.showMessageDialog(null, "Vui lòng nhập số tiền khen thưởng!");
                    return;
                }
                try {
                    // Loại bỏ định dạng tiền tệ và parse
                    String cleaned = soTienText.replaceAll("[^\\d]", "");
                    soTien = Integer.parseInt(cleaned);
                } catch (NumberFormatException ex) {
                    JOptionPane.showMessageDialog(null, "Số tiền không hợp lệ!");
                    return;
                }
            }
        } else {
            // Lấy từ HashMap soTienKT (HashMap bạn khai báo ở đầu class)
            soTien = this.soTienKT.getOrDefault(lyDoKT, 0);
        }
        
        pst.setInt(6, soTien);                                         // Số tiền khen thưởng
        pst.setString(7, lyDoKT);                                      // Lý do khen thưởng
        pst.setString(8, txtMakt.getText().trim());                    // MaKT (WHERE condition)
        
        int result = pst.executeUpdate();
        if (result > 0) {
            JOptionPane.showMessageDialog(null, "Sửa khen thưởng thành công!");
            clearForm();
        } else {
            JOptionPane.showMessageDialog(null, "Sửa khen thưởng thất bại!");
        }
        
    } catch (SQLException e) {
        e.printStackTrace();
        JOptionPane.showMessageDialog(null, "Lỗi sửa khen thưởng: " + e.getMessage());
    }
    
    laydulieu();
}


    // Variables declaration - do not modify                     
    private javax.swing.JButton btnReset;
    private javax.swing.JButton btnSua;
    private javax.swing.JButton btnThem;
    private javax.swing.JButton btnXoa;
    private javax.swing.JButton btnXuat;
    private javax.swing.JComboBox<String> cboLyDoKT;
    private javax.swing.JComboBox<String> cbotennv;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JLabel lbchucvu;
    private javax.swing.JLabel lblkt;
    private javax.swing.JLabel lbmanv;
    private javax.swing.JLabel lbpb;
    private javax.swing.JTextField txtMakt;
    private javax.swing.JTextField txtTimkiem;
    // End of variables declaration                   
}
