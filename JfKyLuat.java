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
public class JfKyLuat extends javax.swing.JPanel {

    private Map<String, Integer> soTienKL = new HashMap<>();

    /**
     * Creates new form JfKhenThuong
     */
    public JfKyLuat() {
        initComponents();
        initComboBoxData();

        cbolydokl.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent evt) {
                cboLyDoKLItemStateChanged(evt);
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
                timkiemkyluat(txtTimkiem.getText().trim());
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

    private void initComboBoxData() {
        soTienKL.put("Đi muộn", 200000);
        soTienKL.put("Nghỉ không phép", 500000);
        soTienKL.put("Vi phạm nội quy", 300000);
        soTienKL.put("Không hoàn thành nhiệm vụ", 400000);
        soTienKL.put("Gây mất đoàn kết", 350000);
        soTienKL.put("Không tuân thủ an toàn", 450000);
        soTienKL.put("Làm việc riêng trong giờ", 250000);
        soTienKL.put("Thái độ không hợp tác", 300000);
        soTienKL.put("Khác", 0); // Cho phép người dùng nhập

        for (String lyDo : soTienKL.keySet()) {
            cbolydokl.addItem(lyDo);
        }
    }

    private void cboLyDoKLItemStateChanged(java.awt.event.ItemEvent evt) {
        if (evt.getStateChange() == ItemEvent.SELECTED) {
            String lyDo = cbolydokl.getSelectedItem().toString();

            if (lyDo.equals("Khác")) {
                String nhap = JOptionPane.showInputDialog(this, "Nhập số tiền kỷ luật:");
                if (nhap != null && !nhap.trim().isEmpty()) {
                    try {
                        String cleaned = nhap.replaceAll("[^\\d]", "");
                        int soTien = Integer.parseInt(cleaned);

                        lblkl.setText(String.format("%,d VNĐ", soTien));
                        lblkl.putClientProperty("rawValue", soTien);
                    } catch (NumberFormatException e) {
                        JOptionPane.showMessageDialog(this, "Số tiền không hợp lệ.", "Lỗi", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    lblkl.setText("");
                }
            } else {
                int soTien = soTienKL.getOrDefault(lyDo, 0);
                lblkl.setText(String.format("%,d VNĐ", soTien));
                lblkl.putClientProperty("rawValue", soTien);
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
        txtMakl = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        cbotennv = new javax.swing.JComboBox<>();
        lbmanv = new javax.swing.JLabel();
        cbolydokl = new javax.swing.JComboBox<>();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        lbpb = new javax.swing.JLabel();
        lbchucvu = new javax.swing.JLabel();
        lblkl = new javax.swing.JLabel();
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

        txtMakl.setFont(new java.awt.Font("SansSerif", 0, 12)); // NOI18N
        txtMakl.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

        jLabel1.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel1.setText("Mã kỷ luật");

        jLabel2.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel2.setText("Mã nhân viên");

        jLabel3.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel3.setText("Tên nhân viên");

        jLabel4.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel4.setText("Kỷ luật");

        jLabel5.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel5.setText("Ngày kỷ luật");

        jLabel6.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel6.setText("Lý do kỷ luật");

        cbotennv.setBorder(new javax.swing.border.MatteBorder(null));

        lbmanv.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbmanv.setText("Mã nhân viên");

        cbolydokl.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "None" }));
        cbolydokl.setBorder(new javax.swing.border.MatteBorder(null));

        jLabel9.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel9.setText("Phòng ban");

        jLabel10.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel10.setText("Chức vụ");

        lbpb.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbpb.setText("Phòng ban");

        lbchucvu.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lbchucvu.setText("Chức vụ");

        lblkl.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        lblkl.setText("Kỷ luật");

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
        jLabel8.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/raven/icon/deadline.png"))); // NOI18N
        jLabel8.setText("Kỷ Luật");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel1)
                            .addComponent(jLabel9))
                        .addGap(40, 40, 40)
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lbchucvu)
                            .addComponent(lbpb)
                            .addComponent(lbmanv)
                            .addComponent(cbotennv, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtMakl, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addContainerGap(65, Short.MAX_VALUE))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel3)
                            .addComponent(jLabel10)
                            .addComponent(jLabel2)
                            .addComponent(jLabel5)
                            .addComponent(jLabel4)
                            .addComponent(jLabel6)
                            .addComponent(btnThem, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel2Layout.createSequentialGroup()
                                        .addGap(13, 13, 13)
                                        .addComponent(btnSua, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(btnXoa, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addComponent(jLabel8))
                                .addGap(46, 46, 46))
                            .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(lblkl)
                                    .addComponent(cbolydokl, javax.swing.GroupLayout.PREFERRED_SIZE, 220, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addContainerGap(15, Short.MAX_VALUE))))))
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(48, 48, 48)
                .addComponent(btnReset, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(54, 54, 54)
                .addComponent(btnXuat, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addGap(41, 41, 41)
                .addComponent(jLabel8)
                .addGap(46, 46, 46)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addGap(39, 39, 39)
                        .addComponent(jLabel3))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(txtMakl, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(32, 32, 32)
                        .addComponent(cbotennv, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(38, 38, 38)
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
                    .addComponent(lblkl))
                .addGap(47, 47, 47)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(cbolydokl, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
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
        txtTimkiem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtTimkiemActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(0, 359, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(txtTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, 400, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 838, javax.swing.GroupLayout.PREFERRED_SIZE)))
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
            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );
    }// </editor-fold>                        

    private void btnThemActionPerformed(java.awt.event.ActionEvent evt) {                                        
        themKyLuat();
    }                                       

    private void btnSuaActionPerformed(java.awt.event.ActionEvent evt) {                                       
        suaKyLuat();
    }                                      

    private void btnXoaActionPerformed(java.awt.event.ActionEvent evt) {                                       

        int selectedRow = jTable1.getSelectedRow();

        if (selectedRow == -1) {
            ThongBao("Vui lòng chọn một Kỷ luật để xóa!", "Lỗi", JOptionPane.ERROR_MESSAGE);
            return;
        }

        String makt = jTable1.getValueAt(selectedRow, 0).toString();

        int confirm = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn muốn xóa kỷ luật này?", "Xác nhận xóa", JOptionPane.YES_NO_OPTION);

        if (confirm == JOptionPane.YES_OPTION) {

            try (Connection conn = connection.getConnection()) {
                String deleteSql = "DELETE FROM Tb_KyLuat WHERE MaKL = ?";
                PreparedStatement psDelete = conn.prepareStatement(deleteSql);
                psDelete.setString(1, makt);

                int rowsAffected = psDelete.executeUpdate();

                if (rowsAffected > 0) {
                    ThongBao("Đã xóa kỷ luật thành công!", "Thông báo", JOptionPane.INFORMATION_MESSAGE);
                    laydulieu();
                } else {
                    ThongBao("Không tìm thấy kỷ luật với mã đã chọn!", "Thông báo", JOptionPane.WARNING_MESSAGE);
                }
            } catch (Exception e) {
                System.out.println(e.toString());
                ThongBao("Có lỗi xảy ra trong quá trình xóa kỷ luật.", "Lỗi", JOptionPane.ERROR_MESSAGE);
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

    private void txtTimkiemActionPerformed(java.awt.event.ActionEvent evt) {                                           
        // TODO add your handling code here:
    }                                          
    private void xuatPDF() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu file PDF");
        fileChooser.setSelectedFile(new File("Kyluat.pdf"));
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
        fileChooser.setSelectedFile(new File("kyluat.xlsx"));
        int userSelection = fileChooser.showSaveDialog(this);

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("DanhSachKyLuat");

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

    private boolean kiemTraMaKLTonTai(String maKL) {
        String sql = "SELECT COUNT(*) FROM Tb_KyLuat WHERE MaKL = ?";

        try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {

            pst.setString(1, maKL);
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
        if (txtMakl.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "Vui lòng nhập mã kỷ luật!");
            txtMakl.requestFocus();
            return false;
        }

        if (cbotennv.getSelectedIndex() == 0) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn nhân viên!");
            cbotennv.requestFocus();
            return false;
        }

        if (DateNgaykl.getDate() == null) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn ngày kỷ luật!");
            DateNgaykl.requestFocus();
            return false;
        }

        String lyDo = (String) cbolydokl.getSelectedItem();
        if (lyDo == null || lyDo.equalsIgnoreCase("None") || lyDo.equals("-- Chọn lý do --") || cbolydokl.getSelectedIndex() == 0) {
            JOptionPane.showMessageDialog(null, "Vui lòng chọn lý do khen thưởng!");
            cbolydokl.requestFocus();
            return false;
        }

        if (lblkl.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "Vui lòng nhập lý do khen thưởng!");
            lblkl.requestFocus();
            return false;
        }

        return true;
    }

    private void timkiemkyluat(String keyword) {
        if (keyword == null || keyword.trim().isEmpty()) {
            laydulieu(); // Gọi lại hàm hiển thị toàn bộ nếu không có từ khóa
            return;
        }

        String sql = "SELECT * FROM Tb_KyLuat WHERE MaKL LIKE ? OR MaNhanVien LIKE ? OR TenNhanVien LIKE ? OR KyLuat LIKE ?";

        try (Connection conn = connection.getConnection(); PreparedStatement ps = conn.prepareStatement(sql)) {
            String searchPattern = "%" + keyword + "%";
            for (int i = 1; i <= 4; i++) {
                ps.setString(i, searchPattern);
            }

            DefaultTableModel tableModel = new DefaultTableModel(
                    new Object[]{"Mã KL", "Mã NV", "Tên nhân viên", "Phòng ban", "Chức vụ", "Ngày KL", "Kỷ luật", "Lý do KL"}, 0
            );
            jTable1.setModel(tableModel);

            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                Object[] item = new Object[8];
                item[0] = rs.getString("MaKL");
                item[1] = rs.getString("MaNhanVien");
                item[2] = rs.getString("TenNhanVien");
                item[3] = rs.getString("PhongBan"); // Vẫn hiển thị mã phòng ban
                item[4] = rs.getString("ChucVu");
                item[5] = rs.getDate("NgayKL");
                item[6] = rs.getString("KyLuat");
                item[7] = rs.getString("LyDoKL");

                tableModel.addRow(item);
            }
        } catch (SQLException e) {
            e.printStackTrace();
            ThongBao("Có lỗi khi tìm kiếm dữ liệu kỷ luật.", "Lỗi", JOptionPane.ERROR_MESSAGE);
        }
    }

    public void clearForm() {
        txtMakl.setText("");
        cbotennv.setSelectedIndex(0);
        lbmanv.setText("");
        lbpb.setText("");
        lbchucvu.setText("");
        DateNgaykl.setDate(null);
        cbolydokl.setSelectedIndex(0);
        lblkl.setText("");
    }

    private void clearThongTinNhanVien() {
        lbmanv.setText("");
        lbpb.setText("");
        lbchucvu.setText("");
        lblkl.setText("");
    }

    private void Nhapdulieu() {
        int selectedRow = jTable1.getSelectedRow();

        if (selectedRow != -1) {

            String maKyLuat = jTable1.getValueAt(selectedRow, 0).toString();
            String tenNhanVien = jTable1.getValueAt(selectedRow, 2).toString();
            String maNhanVien = jTable1.getValueAt(selectedRow, 1).toString();
            String PhongBan = jTable1.getValueAt(selectedRow, 3).toString();
            String ChucVu = jTable1.getValueAt(selectedRow, 4).toString();
            Date ngayKl = (Date) jTable1.getValueAt(selectedRow, 5);
            String kyLuat = jTable1.getValueAt(selectedRow, 6).toString();
            String lyDokl = jTable1.getValueAt(selectedRow, 7).toString();

            txtMakl.setText(maKyLuat);
            cbotennv.setSelectedItem(tenNhanVien);
            lbmanv.setText(maNhanVien);
            lbpb.setText(PhongBan);
            lbchucvu.setText(ChucVu);
            DateNgaykl.setDate(ngayKl);

            cbolydokl.setSelectedItem(lyDokl);

            lblkl.setText(kyLuat);

        }
    }

    private void laydulieu() {
        String sql = "SELECT MaKL, MaNhanVien, TenNhanVien, PhongBan, ChucVu, NgayKL, KyLuat, LyDoKL FROM Tb_KyLuat";

        try (Connection conn = connection.getConnection(); PreparedStatement ps = conn.prepareStatement(sql)) {
            ResultSet rs = ps.executeQuery();

            // Đặt tiêu đề cột đầy đủ
            DefaultTableModel tableModel = new DefaultTableModel(
                    new Object[]{"Mã KL", "Mã NV", "Tên nhân viên", "Phòng ban", "Chức vụ", "Ngày KL", "Kỷ luật", "Lý do KL"}, 0
            );
            jTable1.setModel(tableModel);

            while (rs.next()) {
                Object[] row = new Object[8];
                row[0] = rs.getString("MaKL");
                row[1] = rs.getString("MaNhanVien");
                row[2] = rs.getString("TenNhanVien");
                row[3] = rs.getString("PhongBan");
                row[4] = rs.getString("ChucVu");
                row[5] = rs.getDate("NgayKL");
                row[6] = rs.getString("KyLuat");
                row[7] = rs.getString("LyDoKL");

                tableModel.addRow(row);
            }

        } catch (SQLException e) {
            e.printStackTrace();
            ThongBao("Lỗi khi tải dữ liệu khen thưởng.", "Lỗi", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void themKyLuat() {
        if (!validateInput()) {
            return;
        }
        if (kiemTraMaKLTonTai(txtMakl.getText().trim())) {
            JOptionPane.showMessageDialog(null, "Mã kỷ luật đã tồn tại!");
            return;
        }

        String sql = "INSERT INTO Tb_KyLuat (MaKL, MaNhanVien, TenNhanVien, PhongBan, ChucVu, NgayKL, KyLuat, LyDoKL) VALUES (?, ?, ?, ?, ?, ?, ?, ?)";

        try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {
            pst.setString(1, txtMakl.getText().trim());                     // MaKL
            pst.setString(2, lbmanv.getText().trim());                     // MaNhanVien

            // Tách tên nhân viên từ combo box: "NV01 - Nguyễn Văn A" -> "Nguyễn Văn A"
            String selected = (String) cbotennv.getSelectedItem();
            String tenNhanVien = selected.contains(" - ") ? selected.split(" - ", 2)[1] : selected;
            pst.setString(3, tenNhanVien);                                 // TenNhanVien

            pst.setString(4, lbpb.getText().trim());                       // PhongBan
            pst.setString(5, lbchucvu.getText().trim());                   // ChucVu

            // Ngày kỷ luật
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            String ngayKL = sdf.format(DateNgaykl.getDate());
            pst.setString(6, ngayKL);

            // Lấy lý do kỷ luật
            String lyDoKL = (String) cbolydokl.getSelectedItem();

            // Lấy số tiền kỷ luật
            int soTien;
            if (lyDoKL.equals("Khác")) {
                // Lấy giá trị raw từ clientProperty hoặc parse từ label
                Object rawValue = lblkl.getClientProperty("rawValue");
                if (rawValue != null && rawValue instanceof Integer) {
                    soTien = (Integer) rawValue;
                } else {
                    // Fallback: parse từ text của label
                    String soTienText = lblkl.getText().trim();
                    if (soTienText.isEmpty() || soTienText.equals("")) {
                        JOptionPane.showMessageDialog(null, "Vui lòng nhập số tiền kỷ luật!");
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
                // Lấy từ HashMap soTienKL
                soTien = this.soTienKL.getOrDefault(lyDoKL, 0);
            }

            pst.setInt(7, soTien);                                         // Số tiền kỷ luật
            pst.setString(8, lyDoKL);                                      // Lý do kỷ luật

            int result = pst.executeUpdate();
            if (result > 0) {
                JOptionPane.showMessageDialog(null, "Thêm kỷ luật thành công!");
                clearForm();
            } else {
                JOptionPane.showMessageDialog(null, "Thêm kỷ luật thất bại!");
            }

        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Lỗi thêm kỷ luật: " + e.getMessage());
        }

        laydulieu();
    }
private void suaKyLuat() {
    if (!validateInput()) {
        return;
    }
    
    // Kiểm tra xem có dòng nào được chọn không
    int selectedRow = jTable1.getSelectedRow();
    if (selectedRow == -1) {
        JOptionPane.showMessageDialog(null, "Vui lòng chọn một dòng để sửa!");
        return;
    }
    
    String sql = "UPDATE Tb_KyLuat SET MaNhanVien=?, TenNhanVien=?, PhongBan=?, ChucVu=?, NgayKL=?, KyLuat=?, LyDoKL=? WHERE MaKL=?";
    
    try (Connection conn = connection.getConnection(); PreparedStatement pst = conn.prepareStatement(sql)) {
        pst.setString(1, lbmanv.getText().trim());                     // MaNhanVien
        
        // Tách tên nhân viên từ combo box: "NV01 - Nguyễn Văn A" -> "Nguyễn Văn A"
        String selected = (String) cbotennv.getSelectedItem();
        String tenNhanVien = selected.contains(" - ") ? selected.split(" - ", 2)[1] : selected;
        pst.setString(2, tenNhanVien);                                 // TenNhanVien
        
        pst.setString(3, lbpb.getText().trim());                       // PhongBan
        pst.setString(4, lbchucvu.getText().trim());                   // ChucVu
        
        // Ngày kỷ luật
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String ngayKL = sdf.format(DateNgaykl.getDate());
        pst.setString(5, ngayKL);
        
        // Lấy lý do kỷ luật
        String lyDoKL = (String) cbolydokl.getSelectedItem();
        
        // Lấy số tiền kỷ luật
        int soTien;
        if (lyDoKL.equals("Khác")) {
            // Lấy giá trị raw từ clientProperty hoặc parse từ label
            Object rawValue = lblkl.getClientProperty("rawValue");
            if (rawValue != null && rawValue instanceof Integer) {
                soTien = (Integer) rawValue;
            } else {
                // Fallback: parse từ text của label
                String soTienText = lblkl.getText().trim();
                if (soTienText.isEmpty() || soTienText.equals("")) {
                    JOptionPane.showMessageDialog(null, "Vui lòng nhập số tiền kỷ luật!");
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
            // Lấy từ HashMap soTienKL
            soTien = this.soTienKL.getOrDefault(lyDoKL, 0);
        }
        
        pst.setInt(6, soTien);                                         // Số tiền kỷ luật
        pst.setString(7, lyDoKL);                                      // Lý do kỷ luật
        pst.setString(8, txtMakl.getText().trim());                    // MaKL (WHERE condition)
        
        int result = pst.executeUpdate();
        if (result > 0) {
            JOptionPane.showMessageDialog(null, "Sửa kỷ luật thành công!");
            clearForm();
        } else {
            JOptionPane.showMessageDialog(null, "Sửa kỷ luật thất bại!");
        }
        
    } catch (SQLException e) {
        e.printStackTrace();
        JOptionPane.showMessageDialog(null, "Lỗi sửa kỷ luật: " + e.getMessage());
    }
    
    laydulieu();
}

    // Variables declaration - do not modify                     
    private javax.swing.JButton btnReset;
    private javax.swing.JButton btnSua;
    private javax.swing.JButton btnThem;
    private javax.swing.JButton btnXoa;
    private javax.swing.JButton btnXuat;
    private javax.swing.JComboBox<String> cbolydokl;
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
    private javax.swing.JLabel lblkl;
    private javax.swing.JLabel lbmanv;
    private javax.swing.JLabel lbpb;
    private javax.swing.JTextField txtMakl;
    private javax.swing.JTextField txtTimkiem;
    // End of variables declaration                   
}
