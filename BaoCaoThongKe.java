package com.raven.form;

import javax.swing.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import javax.swing.table.DefaultTableModel;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.*;
import java.util.Vector;
import java.util.HashMap;
import java.util.Map;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Form báo cáo thống kê với JTabbedPane, biểu đồ và giao diện được tối ưu
 */
public class BaoCaoThongKe extends JPanel {

    private JTabbedPane tabbedPane;
    private JPanel panelTongQuat, panelHopDong, panelPhongBan, panelBieuDo;

    // Components cho tab Thống kê tổng quát
    private JLabel lblTongNV, lblTongPB, lblTongHD, lblTongBL;
    private JLabel lblValueTongNV, lblValueTongPB, lblValueTongHD, lblValueTongBL;
    private JTable tableTongQuat;
    private DefaultTableModel modelTongQuat;

    // Components cho tab Hợp đồng
    private JTable tableHopDong;
    private DefaultTableModel modelHopDong;
    private JComboBox<String> cmbLoaiHD;
    private JButton btnLocHD, btnRefreshHD;

    // Components cho tab Phòng ban
    private JTable tablePhongBan;
    private DefaultTableModel modelPhongBan;
    private JComboBox<String> cmbPhongBan;
    private JButton btnLocPB, btnRefreshPB;

    // Components cho tab Biểu đồ
    private JPanel panelChart;
    private JComboBox<String> cmbChartType;
    private JButton btnUpdateChart;

    public BaoCaoThongKe() {
        initComponents();

        // Thêm WindowListener
        setupLayout();
        loadData();
    }

    private void initComponents() {
        setLayout(new BorderLayout());
        setSize(1197, 762);

        // Đặt Look and Feel hiện đại
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Khởi tạo JTabbedPane với style đẹp hơn
        tabbedPane = new JTabbedPane();
        tabbedPane.setFont(new Font("Segoe UI", Font.BOLD, 12));
        tabbedPane.setBackground(new Color(245, 245, 245));

        // Khởi tạo các panel
        initTongQuatPanel();
        initHopDongPanel();
        initPhongBanPanel();
        initBieuDoPanel();

        // Thêm các tab với icon
        tabbedPane.addTab("Thống Kê Tổng Quát", panelTongQuat);
        tabbedPane.addTab("Báo Cáo Hợp Đồng", panelHopDong);
        tabbedPane.addTab("Báo Cáo Phòng Ban", panelPhongBan);
        tabbedPane.addTab("Biểu Đồ", panelBieuDo);
    }

    private void initTongQuatPanel() {
        panelTongQuat = new JPanel(new BorderLayout(10, 10));
        panelTongQuat.setBackground(Color.WHITE);

        // Panel hiển thị số liệu tổng quan với design card
        JPanel panelSoLieu = new JPanel(new GridLayout(2, 4, 15, 15));
        panelSoLieu.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder(BorderFactory.createLineBorder(new Color(70, 130, 180), 2),
                        "Thống Kê Tổng Quan", 0, 0, new Font("Segoe UI", Font.BOLD, 14), new Color(70, 130, 180)),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        panelSoLieu.setBackground(Color.WHITE);

        // Tạo các card thống kê với màu sắc đẹp
        lblTongNV = createStatsLabel("Tổng Nhân Viên:");
        lblValueTongNV = createStatsValueLabel("0", new Color(52, 152, 219));

        lblTongPB = createStatsLabel("Tổng Phòng Ban:");
        lblValueTongPB = createStatsValueLabel("0", new Color(46, 204, 113));

        lblTongHD = createStatsLabel("Tổng Hợp Đồng:");
        lblValueTongHD = createStatsValueLabel("0", new Color(231, 76, 60));

        lblTongBL = createStatsLabel("Tổng Bảng Lương:");
        lblValueTongBL = createStatsValueLabel("0", new Color(243, 156, 18));

        panelSoLieu.add(lblTongNV);
        panelSoLieu.add(lblTongPB);
        panelSoLieu.add(lblTongHD);
        panelSoLieu.add(lblTongBL);
        panelSoLieu.add(lblValueTongNV);
        panelSoLieu.add(lblValueTongPB);
        panelSoLieu.add(lblValueTongHD);
        panelSoLieu.add(lblValueTongBL);

        // Panel nút xuất Excel với style đẹp
        JPanel panelButtonTQ = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 10));
        panelButtonTQ.setBackground(Color.WHITE);

        JButton btnXuatTongQuat = createStyledButton("Xuất Excel - Thống Kê", new Color(52, 152, 219));
        JButton btnBaoCaoTongHop = createStyledButton("Tạo Báo Cáo Tổng Hợp", new Color(46, 204, 113));

        btnXuatTongQuat.addActionListener(e -> xuatExcelTongQuat());
        btnBaoCaoTongHop.addActionListener(e -> taoBaoCaoTongHop());

        panelButtonTQ.add(btnXuatTongQuat);
        panelButtonTQ.add(btnBaoCaoTongHop);

        // Bảng thống kê chi tiết với style đẹp
        String[] columnsTongQuat = {"Phòng Ban", "Số NV", "Số Hợp Đồng", "Lương TB", "Khen Thưởng", "Kỷ Luật"};
        modelTongQuat = new DefaultTableModel(columnsTongQuat, 0);
        tableTongQuat = new JTable(modelTongQuat);
        styleTable(tableTongQuat);

        JScrollPane scrollTongQuat = new JScrollPane(tableTongQuat);
        scrollTongQuat.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder(BorderFactory.createLineBorder(new Color(70, 130, 180), 2),
                        "Chi Tiết Theo Phòng Ban", 0, 0, new Font("Segoe UI", Font.BOLD, 12), new Color(70, 130, 180)),
                BorderFactory.createEmptyBorder(5, 5, 5, 5)
        ));

        // Panel chứa button và bảng
        JPanel panelContent = new JPanel(new BorderLayout(10, 10));
        panelContent.setBackground(Color.WHITE);
        panelContent.add(panelButtonTQ, BorderLayout.NORTH);
        panelContent.add(scrollTongQuat, BorderLayout.CENTER);

        panelTongQuat.add(panelSoLieu, BorderLayout.NORTH);
        panelTongQuat.add(panelContent, BorderLayout.CENTER);
    }

    private void initHopDongPanel() {
        panelHopDong = new JPanel(new BorderLayout(10, 10));
        panelHopDong.setBackground(Color.WHITE);

        // Panel điều khiển với style đẹp
        JPanel panelControl = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        panelControl.setBackground(new Color(248, 249, 250));
        panelControl.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(200, 200, 200)),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));

        cmbLoaiHD = new JComboBox<>(new String[]{
            "Tất cả",
            "Hợp đồng không xác định thời hạn",
            "6 tháng", "3 tháng", "5 tháng",
            "1 năm", "2 năm", "3 năm", "4 năm", "5 năm",
            "Hợp đồng làm việc bán thời gian",
            "Hợp đồng cộng tác viên",
            "Hợp đồng thử việc"
        });
        styleComboBox(cmbLoaiHD);

        btnLocHD = createStyledButton("Lọc", new Color(52, 152, 219));
        btnRefreshHD = createStyledButton("Làm mới", new Color(95, 39, 205));
        JButton btnXuatHopDong = createStyledButton("Xuất Excel", new Color(46, 204, 113));

        JLabel lblLoaiHD = new JLabel("Loại hợp đồng:");
        lblLoaiHD.setFont(new Font("Segoe UI", Font.BOLD, 12));

        panelControl.add(lblLoaiHD);
        panelControl.add(cmbLoaiHD);
        panelControl.add(btnLocHD);
        panelControl.add(btnRefreshHD);
        panelControl.add(btnXuatHopDong);

        // Bảng hợp đồng
        String[] columnsHD = {"Mã HD", "Tên NV", "Loại HD", "Ngày ký", "Ngày hết hạn", "Phòng ban", "Lương"};
        modelHopDong = new DefaultTableModel(columnsHD, 0);
        tableHopDong = new JTable(modelHopDong);
        styleTable(tableHopDong);

        JScrollPane scrollHD = new JScrollPane(tableHopDong);
        scrollHD.setBorder(BorderFactory.createLineBorder(new Color(200, 200, 200)));

        // Sự kiện
        btnLocHD.addActionListener(e -> loadHopDongData());
        btnRefreshHD.addActionListener(e -> {
            cmbLoaiHD.setSelectedIndex(0);
            loadHopDongData();
        });
        btnXuatHopDong.addActionListener(e -> xuatExcelHopDong());

        panelHopDong.add(panelControl, BorderLayout.NORTH);
        panelHopDong.add(scrollHD, BorderLayout.CENTER);
    }

    private void initPhongBanPanel() {
        panelPhongBan = new JPanel(new BorderLayout(10, 10));
        panelPhongBan.setBackground(Color.WHITE);

        // Panel điều khiển
        JPanel panelControl = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        panelControl.setBackground(new Color(248, 249, 250));
        panelControl.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(200, 200, 200)),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));

        cmbPhongBan = new JComboBox<>();
        styleComboBox(cmbPhongBan);

        btnLocPB = createStyledButton("Lọc", new Color(52, 152, 219));
        btnRefreshPB = createStyledButton("Làm mới", new Color(95, 39, 205));
        JButton btnXuatPhongBan = createStyledButton("Xuất Excel", new Color(46, 204, 113));

        JLabel lblPhongBan = new JLabel("Phòng ban:");
        lblPhongBan.setFont(new Font("Segoe UI", Font.BOLD, 12));

        panelControl.add(lblPhongBan);
        panelControl.add(cmbPhongBan);
        panelControl.add(btnLocPB);
        panelControl.add(btnRefreshPB);
        panelControl.add(btnXuatPhongBan);

        // Bảng nhân viên theo phòng ban
        String[] columnsPB = {"Mã NV", "Tên NV", "Chức vụ", "Ngày vào làm", "Lương CB", "Loại HD"};
        modelPhongBan = new DefaultTableModel(columnsPB, 0);
        tablePhongBan = new JTable(modelPhongBan);
        styleTable(tablePhongBan);

        JScrollPane scrollPB = new JScrollPane(tablePhongBan);
        scrollPB.setBorder(BorderFactory.createLineBorder(new Color(200, 200, 200)));

        // Sự kiện
        btnLocPB.addActionListener(e -> loadPhongBanData());
        btnRefreshPB.addActionListener(e -> loadPhongBanData());
        btnXuatPhongBan.addActionListener(e -> xuatExcelPhongBan());

        panelPhongBan.add(panelControl, BorderLayout.NORTH);
        panelPhongBan.add(scrollPB, BorderLayout.CENTER);
    }

    private void initBieuDoPanel() {
        panelBieuDo = new JPanel(new BorderLayout(10, 10));
        panelBieuDo.setBackground(Color.WHITE);

        // Panel điều khiển biểu đồ
        JPanel panelChartControl = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        panelChartControl.setBackground(new Color(248, 249, 250));
        panelChartControl.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(200, 200, 200)),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));

        cmbChartType = new JComboBox<>(new String[]{
            "Biểu đồ nhân viên theo phòng ban",
            "Biểu đồ hợp đồng theo loại",
            "Biểu đồ lương trung bình",
            "Biểu đồ khen thưởng - kỷ luật"
        });
        styleComboBox(cmbChartType);

        btnUpdateChart = createStyledButton("Cập nhật biểu đồ", new Color(155, 89, 182));

        JLabel lblChartType = new JLabel("Loại biểu đồ:");
        lblChartType.setFont(new Font("Segoe UI", Font.BOLD, 12));

        panelChartControl.add(lblChartType);
        panelChartControl.add(cmbChartType);
        panelChartControl.add(btnUpdateChart);

        // Panel hiển thị biểu đồ
        panelChart = new JPanel();
        panelChart.setBackground(Color.WHITE);
        panelChart.setBorder(BorderFactory.createLineBorder(new Color(200, 200, 200)));
        panelChart.setLayout(new BorderLayout());

        btnUpdateChart.addActionListener(e -> updateChart());

        panelBieuDo.add(panelChartControl, BorderLayout.NORTH);
        panelBieuDo.add(panelChart, BorderLayout.CENTER);

        // Load biểu đồ mặc định
        updateChart();
    }

    // Phương thức tạo label thống kê
    private JLabel createStatsLabel(String text) {
        JLabel label = new JLabel(text, JLabel.CENTER);
        label.setFont(new Font("Segoe UI", Font.BOLD, 12));
        label.setForeground(new Color(52, 73, 94));
        return label;
    }

    // Phương thức tạo label giá trị thống kê
    private JLabel createStatsValueLabel(String text, Color color) {
        JLabel label = new JLabel(text, JLabel.CENTER);
        label.setFont(new Font("Segoe UI", Font.BOLD, 20));
        label.setForeground(color);
        label.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(color, 2),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        label.setOpaque(true);
        label.setBackground(new Color(color.getRed(), color.getGreen(), color.getBlue(), 30));
        return label;
    }

    // Phương thức tạo nút có style đẹp
    private JButton createStyledButton(String text, Color color) {
        JButton button = new JButton(text);
        button.setFont(new Font("Segoe UI", Font.BOLD, 11));
        button.setForeground(Color.BLACK);
        button.setBackground(color);
        button.setBorder(BorderFactory.createEmptyBorder(8, 15, 8, 15));
        button.setFocusPainted(false);
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));

        // Hiệu ứng hover
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(color.darker());
            }

            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(color);
            }
        });

        return button;
    }

    // Phương thức style cho ComboBox
    private void styleComboBox(JComboBox<?> comboBox) {
        comboBox.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        comboBox.setPreferredSize(new Dimension(200, 30));
        comboBox.setBorder(BorderFactory.createLineBorder(new Color(200, 200, 200)));
    }

    // Phương thức style cho Table
    private void styleTable(JTable table) {
        table.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        table.setRowHeight(25);
        table.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 11));
        table.getTableHeader().setBackground(new Color(52, 152, 219));
        table.getTableHeader().setForeground(Color.BLACK);
        table.setSelectionBackground(new Color(52, 152, 219, 100));
        table.setGridColor(new Color(220, 220, 220));
        table.setShowVerticalLines(true);
        table.setShowHorizontalLines(true);
    }

    // Phương thức cập nhật biểu đồ
    private void updateChart() {
        String chartType = (String) cmbChartType.getSelectedItem();
        panelChart.removeAll();

        JPanel chartPanel = createChart(chartType);
        panelChart.add(chartPanel, BorderLayout.CENTER);

        panelChart.revalidate();
        panelChart.repaint();
    }

    // Phương thức tạo biểu đồ đơn giản
    private JPanel createChart(String chartType) {
        JPanel chart = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Graphics2D g2d = (Graphics2D) g;
                g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

                try {
                    switch (chartType) {
                        case "Biểu đồ nhân viên theo phòng ban":
                            drawEmployeeChart(g2d);
                            break;
                        case "Biểu đồ hợp đồng theo loại":
                            drawContractChart(g2d);
                            break;
                        case "Biểu đồ lương trung bình":
                            drawSalaryChart(g2d);
                            break;
                        case "Biểu đồ khen thưởng - kỷ luật":
                            drawRewardPunishmentChart(g2d);
                            break;
                    }
                } catch (Exception e) {
                    g2d.setColor(Color.RED);
                    g2d.setFont(new Font("Segoe UI", Font.BOLD, 16));
                    g2d.drawString("Lỗi khi tải dữ liệu biểu đồ", 50, 100);
                }
            }
        };

        chart.setBackground(Color.WHITE);
        chart.setPreferredSize(new Dimension(800, 400));
        return chart;
    }

    // Phương thức vẽ biểu đồ nhân viên theo phòng ban
    private void drawEmployeeChart(Graphics2D g2d) {
        Map<String, Integer> data = getEmployeeByDepartment();
        drawBarChart(g2d, data, "Số lượng nhân viên theo phòng ban", new Color(52, 152, 219));
    }

    // Phương thức vẽ biểu đồ hợp đồng theo loại
    private void drawContractChart(Graphics2D g2d) {
        Map<String, Integer> data = getContractByType();
        drawPieChart(g2d, data, "Phân bố hợp đồng theo loại");
    }

    // Phương thức vẽ biểu đồ lương trung bình
    private void drawSalaryChart(Graphics2D g2d) {
        Map<String, Integer> data = getAverageSalaryByDepartment();
        drawBarChart(g2d, data, "Lương trung bình theo phòng ban (triệu VNĐ)", new Color(46, 204, 113));
    }

    // Phương thức vẽ biểu đồ khen thưởng - kỷ luật
    private void drawRewardPunishmentChart(Graphics2D g2d) {
        Map<String, Integer> rewards = getRewardsByDepartment();
        Map<String, Integer> punishments = getPunishmentsByDepartment();
        drawComparisonChart(g2d, rewards, punishments, "Khen thưởng vs Kỷ luật theo phòng ban");
    }

    // Phương thức vẽ biểu đồ cột
    private void drawBarChart(Graphics2D g2d, Map<String, Integer> data, String title, Color color) {
        if (data.isEmpty()) {
            return;
        }

        int width = panelChart.getWidth() - 100;
        int height = panelChart.getHeight() - 150;
        int x = 50;
        int y = 50;

        // Vẽ tiêu đề
        g2d.setColor(Color.BLACK);
        g2d.setFont(new Font("Segoe UI", Font.BOLD, 16));
        FontMetrics fm = g2d.getFontMetrics();
        int titleWidth = fm.stringWidth(title);
        g2d.drawString(title, (panelChart.getWidth() - titleWidth) / 2, 30);

        // Vẽ trục
        g2d.drawLine(x, y + height, x + width, y + height); // Trục X
        g2d.drawLine(x, y, x, y + height); // Trục Y

        // Tìm giá trị max
        int maxValue = data.values().stream().mapToInt(Integer::intValue).max().orElse(1);

        // Vẽ các cột
        int barWidth = width / data.size() - 10;
        int currentX = x + 5;

        for (Map.Entry<String, Integer> entry : data.entrySet()) {
            int barHeight = (int) ((double) entry.getValue() / maxValue * height * 0.8);

            // Vẽ cột
            g2d.setColor(color);
            g2d.fillRect(currentX, y + height - barHeight, barWidth, barHeight);

            // Vẽ viền cột
            g2d.setColor(color.darker());
            g2d.drawRect(currentX, y + height - barHeight, barWidth, barHeight);

            // Vẽ giá trị
            g2d.setColor(Color.BLACK);
            g2d.setFont(new Font("Segoe UI", Font.PLAIN, 10));
            String value = String.valueOf(entry.getValue());
            fm = g2d.getFontMetrics();
            int valueWidth = fm.stringWidth(value);
            g2d.drawString(value, currentX + (barWidth - valueWidth) / 2, y + height - barHeight - 5);

            // Vẽ tên phòng ban
            String label = entry.getKey();
            if (label.length() > 50) {
                label = label.substring(0, 50) + "...";
            }
            int labelWidth = fm.stringWidth(label);
            g2d.drawString(label, currentX + (barWidth - labelWidth) / 2, y + height + 15);

            currentX += barWidth + 10;
        }
    }

    // Phương thức vẽ biểu đồ tròn
    private void drawPieChart(Graphics2D g2d, Map<String, Integer> data, String title) {
        if (data.isEmpty()) {
            return;
        }

        int centerX = panelChart.getWidth() / 2;
        int centerY = panelChart.getHeight() / 2 + 20;
        int radius = Math.min(panelChart.getWidth(), panelChart.getHeight()) / 4;

        // Vẽ tiêu đề
        g2d.setColor(Color.BLACK);
        g2d.setFont(new Font("Segoe UI", Font.BOLD, 16));
        FontMetrics fm = g2d.getFontMetrics();
        int titleWidth = fm.stringWidth(title);
        g2d.drawString(title, (panelChart.getWidth() - titleWidth) / 2, 30);

        // Tính tổng
        int total = data.values().stream().mapToInt(Integer::intValue).sum();
        if (total == 0) {
            return;
        }

        // Màu sắc cho từng phần
        Color[] colors = {
            new Color(52, 152, 219), new Color(46, 204, 113), new Color(231, 76, 60),
            new Color(243, 156, 18), new Color(155, 89, 182), new Color(230, 126, 34)
        };

        int startAngle = 0;
        int colorIndex = 0;

        for (Map.Entry<String, Integer> entry : data.entrySet()) {
            int angle = (int) ((double) entry.getValue() / total * 360);

            // Vẽ phần của biểu đồ tròn
            g2d.setColor(colors[colorIndex % colors.length]);
            g2d.fillArc(centerX - radius, centerY - radius, radius * 2, radius * 2, startAngle, angle);

            // Vẽ viền
            g2d.setColor(Color.WHITE);
            g2d.setStroke(new BasicStroke(2));
            g2d.drawArc(centerX - radius, centerY - radius, radius * 2, radius * 2, startAngle, angle);

            startAngle += angle;
            colorIndex++;
        }

        // Vẽ chú thích
        int legendY = centerY + radius + 30;
        colorIndex = 0;
        for (Map.Entry<String, Integer> entry : data.// Tiếp tục phần vẽ chú thích của drawPieChart
                entrySet()) {
            // Vẽ ô màu
            g2d.setColor(colors[colorIndex % colors.length]);
            g2d.fillRect(50, legendY, 15, 15);
            g2d.setColor(Color.BLACK);
            g2d.drawRect(50, legendY, 15, 15);

            // Vẽ text
            g2d.setFont(new Font("Segoe UI", Font.PLAIN, 11));
            String legendText = entry.getKey() + ": " + entry.getValue() + " ("
                    + String.format("%.1f%%", (double) entry.getValue() / total * 100) + ")";
            g2d.drawString(legendText, 75, legendY + 12);

            legendY += 25;
            colorIndex++;
        }
    }

    // Phương thức vẽ biểu đồ so sánh
    private void drawComparisonChart(Graphics2D g2d, Map<String, Integer> rewards,
            Map<String, Integer> punishments, String title) {
        if (rewards.isEmpty() && punishments.isEmpty()) {
            return;
        }

        int width = panelChart.getWidth() - 100;
        int height = panelChart.getHeight() - 150;
        int x = 50;
        int y = 50;

        // Vẽ tiêu đề
        g2d.setColor(Color.BLACK);
        g2d.setFont(new Font("Segoe UI", Font.BOLD, 16));
        FontMetrics fm = g2d.getFontMetrics();
        int titleWidth = fm.stringWidth(title);
        g2d.drawString(title, (panelChart.getWidth() - titleWidth) / 2, 30);

        // Vẽ trục
        g2d.drawLine(x, y + height, x + width, y + height); // Trục X
        g2d.drawLine(x, y, x, y + height); // Trục Y

        // Tìm giá trị max
        int maxReward = rewards.values().stream().mapToInt(Integer::intValue).max().orElse(0);
        int maxPunishment = punishments.values().stream().mapToInt(Integer::intValue).max().orElse(0);
        int maxValue = Math.max(maxReward, maxPunishment);
        if (maxValue == 0) {
            maxValue = 1;
        }

        // Lấy tất cả phòng ban
        java.util.Set<String> allDepts = new java.util.HashSet<>();
        allDepts.addAll(rewards.keySet());
        allDepts.addAll(punishments.keySet());

        // Vẽ các cột
        int groupWidth = width / allDepts.size() - 20;
        int barWidth = groupWidth / 2 - 5;
        int currentX = x + 10;

        for (String dept : allDepts) {
            int rewardValue = rewards.getOrDefault(dept, 0);
            int punishmentValue = punishments.getOrDefault(dept, 0);

            int rewardHeight = (int) ((double) rewardValue / maxValue * height * 0.8);
            int punishmentHeight = (int) ((double) punishmentValue / maxValue * height * 0.8);

            // Vẽ cột khen thưởng (màu xanh lá)
            g2d.setColor(new Color(46, 204, 113));
            g2d.fillRect(currentX, y + height - rewardHeight, barWidth, rewardHeight);
            g2d.setColor(new Color(39, 174, 96));
            g2d.drawRect(currentX, y + height - rewardHeight, barWidth, rewardHeight);

            // Vẽ cột kỷ luật (màu đỏ)
            g2d.setColor(new Color(231, 76, 60));
            g2d.fillRect(currentX + barWidth + 5, y + height - punishmentHeight, barWidth, punishmentHeight);
            g2d.setColor(new Color(192, 57, 43));
            g2d.drawRect(currentX + barWidth + 5, y + height - punishmentHeight, barWidth, punishmentHeight);

            // Vẽ giá trị
            g2d.setColor(Color.BLACK);
            g2d.setFont(new Font("Segoe UI", Font.PLAIN, 9));
            if (rewardValue > 0) {
                g2d.drawString(String.valueOf(rewardValue), currentX + barWidth / 2 - 5, y + height - rewardHeight - 3);
            }
            if (punishmentValue > 0) {
                g2d.drawString(String.valueOf(punishmentValue), currentX + barWidth + 5 + barWidth / 2 - 5, y + height - punishmentHeight - 3);
            }

            // Vẽ tên phòng ban
            String label = dept.length() > 50 ? dept.substring(0, 50) + "..." : dept;
            FontMetrics labelFm = g2d.getFontMetrics();
            int labelWidth = labelFm.stringWidth(label);
            g2d.drawString(label, currentX + (groupWidth - labelWidth) / 2, y + height + 15);

            currentX += groupWidth + 20;
        }

        // Vẽ chú thích
        g2d.setColor(new Color(46, 204, 113));
        g2d.fillRect(x + width - 200, y + 20, 15, 15);
        g2d.setColor(Color.BLACK);
        g2d.drawRect(x + width - 200, y + 20, 15, 15);
        g2d.drawString("Khen thưởng", x + width - 180, y + 32);

        g2d.setColor(new Color(231, 76, 60));
        g2d.fillRect(x + width - 200, y + 40, 15, 15);
        g2d.setColor(Color.BLACK);
        g2d.drawRect(x + width - 200, y + 40, 15, 15);
        g2d.drawString("Kỷ luật", x + width - 180, y + 52);
    }

    // Các phương thức lấy dữ liệu từ database
    private Map<String, Integer> getEmployeeByDepartment() {
        Map<String, Integer> data = new HashMap<>();
        String sql = "SELECT PhongBan, COUNT(*) AS SoLuong "
                + "FROM Tb_NhanVien "
                + "GROUP BY PhongBan "
                + "ORDER BY SoLuong DESC";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                data.put(rs.getString("PhongBan"), rs.getInt("SoLuong"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return data;
    }

    private Map<String, Integer> getContractByType() {
        Map<String, Integer> data = new HashMap<>();
        String sql = "SELECT LoaiHD, COUNT(*) AS SoLuong "
                + "FROM Tb_NhanVien "
                + "GROUP BY LoaiHD "
                + "ORDER BY SoLuong DESC";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                data.put(rs.getString("LoaiHD"), rs.getInt("SoLuong"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return data;
    }

    private Map<String, Integer> getAverageSalaryByDepartment() {
        Map<String, Integer> data = new HashMap<>();
        String sql = "SELECT nv.PhongBan, AVG(bl.LuongCoBan) AS LuongTB "
                + "FROM Tb_NhanVien nv "
                + "LEFT JOIN Tb_BangLuong bl ON nv.MaNhanVien = bl.MaNhanVien "
                + "WHERE bl.LuongCoBan IS NOT NULL "
                + "GROUP BY nv.PhongBan";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                data.put(rs.getString("PhongBan"),
                        (int) (rs.getDouble("LuongTB") / 1000000)); // chia về triệu
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return data;
    }

    private Map<String, Integer> getRewardsByDepartment() {
        Map<String, Integer> data = new HashMap<>();
        String sql = "SELECT PhongBan, COUNT(MaKT) AS SoLuong "
                + "FROM Tb_KhenThuong "
                + "GROUP BY PhongBan";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                data.put(rs.getString("PhongBan"), rs.getInt("SoLuong"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return data;
    }

    private Map<String, Integer> getPunishmentsByDepartment() {
        Map<String, Integer> data = new HashMap<>();
        String sql = "SELECT PhongBan, COUNT(MaKL) AS SoLuong "
                + "FROM Tb_KyLuat "
                + "GROUP BY PhongBan";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                data.put(rs.getString("PhongBan"), rs.getInt("SoLuong"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return data;
    }

    private void setupLayout() {
        setLayout(new BorderLayout());
        add(tabbedPane, BorderLayout.CENTER);

        //Thêm thanh trạng thái
    }

    private void loadData() {
        loadTongQuatData();
        loadHopDongData();
        loadPhongBanData();
        loadPhongBanComboBox();
    }

    private void loadTongQuatData() {
        try (Connection conn = connection.getConnection()) {
            // Load số liệu tổng quan
            String sqlTongNV = "SELECT COUNT(*) as total FROM Tb_NhanVien";
            String sqlTongPB = "SELECT COUNT(*) as total FROM Tb_PhongBan";
            String sqlTongHD = "SELECT COUNT(*) as total FROM Tb_HopDong";
            String sqlTongBL = "SELECT COUNT(*) as total FROM Tb_BangLuong";

            try (PreparedStatement pstmt = conn.prepareStatement(sqlTongNV); ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    lblValueTongNV.setText(String.valueOf(rs.getInt("total")));
                }
            }

            try (PreparedStatement pstmt = conn.prepareStatement(sqlTongPB); ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    lblValueTongPB.setText(String.valueOf(rs.getInt("total")));
                }
            }

            try (PreparedStatement pstmt = conn.prepareStatement(sqlTongHD); ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    lblValueTongHD.setText(String.valueOf(rs.getInt("total")));
                }
            }

            try (PreparedStatement pstmt = conn.prepareStatement(sqlTongBL); ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    lblValueTongBL.setText(String.valueOf(rs.getInt("total")));
                }
            }

            // Load bảng thống kê chi tiết
            loadChiTietThongKe();

        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Lỗi khi tải dữ liệu: " + e.getMessage());
        }
    }

    private void loadChiTietThongKe() {
        modelTongQuat.setRowCount(0);
        String sql = "SELECT pb.TenPB, "
                + "COUNT(DISTINCT nv.MaNhanVien) AS SoNV, "
                + "COUNT(DISTINCT hd.MaHD) AS SoHD, "
                + "COALESCE(AVG(bl.LuongCoBan), 0) AS LuongTB, "
                + "COUNT(DISTINCT kt.MaKT) AS SoKT, "
                + "COUNT(DISTINCT kl.MaKL) AS SoKL "
                + "FROM Tb_PhongBan pb "
                + "LEFT JOIN Tb_NhanVien nv ON pb.TenPB = nv.PhongBan "
                + // Join theo tên phòng ban
                "LEFT JOIN Tb_HopDong hd ON nv.MaNhanVien = hd.MaNhanVien "
                + "LEFT JOIN Tb_BangLuong bl ON nv.MaNhanVien = bl.MaNhanVien "
                + "LEFT JOIN Tb_KhenThuong kt ON nv.MaNhanVien = kt.MaNhanVien "
                + "LEFT JOIN Tb_KyLuat kl ON nv.MaNhanVien = kl.MaNhanVien "
                + "GROUP BY pb.MaPB, pb.TenPB "
                + "ORDER BY pb.TenPB";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                Vector<Object> row = new Vector<>();
                row.add(rs.getString("TenPB"));                       // Tên phòng ban
                row.add(rs.getInt("SoNV"));                            // Số nhân viên
                row.add(rs.getInt("SoHD"));                            // Số hợp đồng
                row.add(String.format("%.0f", rs.getDouble("LuongTB")));  // Lương trung bình
                row.add(rs.getInt("SoKT"));                            // Số khen thưởng
                row.add(rs.getInt("SoKL"));                            // Số kỷ luật
                modelTongQuat.addRow(row);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void loadHopDongData() {
        modelHopDong.setRowCount(0);
        String loaiHD = (String) cmbLoaiHD.getSelectedItem();

        // Câu lệnh SQL đã sửa tên cột và quan hệ đúng với cấu trúc bảng bạn gửi
        String sql = "SELECT hd.MaHD, nv.TenNhanVien, nv.LoaiHD, hd.NgayKy, "
                + "hd.NgayHetHan, nv.PhongBan, hd.LuongCoBan "
                + "FROM Tb_HopDong hd "
                + "JOIN Tb_NhanVien nv ON hd.MaNhanVien = nv.MaNhanVien "
                + "WHERE 1=1";

        // Nếu có lọc theo loại hợp đồng
        if (!"Tất cả".equals(loaiHD)) {
            sql += " AND nv.LoaiHD = ?";
        }

        sql += " ORDER BY hd.NgayKy DESC";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql)) {

            if (!"Tất cả".equals(loaiHD)) {
                pstmt.setString(1, loaiHD);
            }

            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    Vector<Object> row = new Vector<>();
                    row.add(rs.getString("MaHD"));
                    row.add(rs.getString("TenNhanVien"));
                    row.add(rs.getString("LoaiHD"));
                    row.add(rs.getDate("NgayKy"));
                    row.add(rs.getDate("NgayHetHan"));
                    row.add(rs.getString("PhongBan"));
                    row.add(rs.getInt("LuongCoBan"));
                    modelHopDong.addRow(row);
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void loadPhongBanData() {
        modelPhongBan.setRowCount(0);
        String phongBan = (String) cmbPhongBan.getSelectedItem();

        String sql = "SELECT nv.MaNhanVien, nv.TenNhanVien, nv.ChucVu, nv.NgayVaoLam, "
                + "bl.LuongCoBan, nv.LoaiHD "
                + "FROM Tb_NhanVien nv "
                + "LEFT JOIN Tb_BangLuong bl ON nv.MaNhanVien = bl.MaNhanVien "
                + "WHERE 1=1";

        if (phongBan != null && !"Tất cả".equals(phongBan)) {
            sql += " AND nv.PhongBan = ?";
        }

        sql += " ORDER BY nv.TenNhanVien";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql)) {

            if (phongBan != null && !"Tất cả".equals(phongBan)) {
                pstmt.setString(1, phongBan);
            }

            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    Vector<Object> row = new Vector<>();
                    row.add(rs.getString("MaNhanVien"));
                    row.add(rs.getString("TenNhanVien"));
                    row.add(rs.getString("ChucVu"));
                    row.add(rs.getDate("NgayVaoLam"));
                    row.add(rs.getInt("LuongCoBan"));
                    row.add(rs.getString("LoaiHD"));
                    modelPhongBan.addRow(row);
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void loadPhongBanComboBox() {
        cmbPhongBan.removeAllItems();
        cmbPhongBan.addItem("Tất cả");

        String sql = "SELECT TenPB FROM Tb_PhongBan ORDER BY TenPB";
        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                cmbPhongBan.addItem(rs.getString("TenPB"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    // Các phương thức xuất Excel
    private void xuatExcelTongQuat() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu báo cáo thống kê tổng quát");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        fileChooser.setSelectedFile(new File("BaoCaoThongKeTongQuat.xlsx"));

        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();

            // Đảm bảo file có đuôi .xlsx
            if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try {
                // Tạo workbook mới
                Workbook workbook = new XSSFWorkbook();

                // Tạo sheet cho thống kê tổng quan
                Sheet sheetTongQuan = workbook.createSheet("Thống Kê Tổng Quan");
                createTongQuanSheet(workbook, sheetTongQuan);

                // Tạo sheet cho chi tiết theo phòng ban
                Sheet sheetChiTiet = workbook.createSheet("Chi Tiết Phòng Ban");
                createChiTietSheet(workbook, sheetChiTiet);

                // Ghi file
                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                    workbook.write(outputStream);
                }

                workbook.close();

                JOptionPane.showMessageDialog(this,
                        "Xuất Excel thành công!\nFile đã được lưu tại: " + file.getAbsolutePath(),
                        "Thành công", JOptionPane.INFORMATION_MESSAGE);

            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this,
                        "Lỗi khi xuất Excel: " + e.getMessage(),
                        "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void createTongQuanSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề chính
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("BÁO CÁO THỐNG KÊ TỔNG QUÁT");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        // Ngày tạo báo cáo
        rowNum++;
        Row dateRow = sheet.createRow(rowNum++);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Ngày tạo báo cáo: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss")));
        dateCell.setCellStyle(dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

        rowNum++;

        // Phần thống kê tổng quan
        Row sectionRow = sheet.createRow(rowNum++);
        Cell sectionCell = sectionRow.createCell(0);
        sectionCell.setCellValue("I. THỐNG KÊ TỔNG QUAN");
        sectionCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

        rowNum++;

        // Tạo bảng thống kê số liệu
        String[] labels = {"Tổng số nhân viên:", "Tổng số phòng ban:", "Tổng số hợp đồng:", "Tổng số bảng lương:"};
        String[] values = {
            lblValueTongNV.getText(),
            lblValueTongPB.getText(),
            lblValueTongHD.getText(),
            lblValueTongBL.getText()
        };

        for (int i = 0; i < labels.length; i++) {
            Row dataRow = sheet.createRow(rowNum++);

            Cell labelCell = dataRow.createCell(0);
            labelCell.setCellValue(labels[i]);
            labelCell.setCellStyle(dataStyle);

            Cell valueCell = dataRow.createCell(1);
            try {
                valueCell.setCellValue(Double.parseDouble(values[i]));
            } catch (NumberFormatException e) {
                valueCell.setCellValue(values[i]);
            }
            valueCell.setCellStyle(numberStyle);
        }

        // Điều chỉnh độ rộng cột
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 3000);
    }

    private void createChiTietSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("II. CHI TIẾT THEO PHÒNG BAN");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        rowNum++;

        // Header của bảng
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Phòng Ban", "Số NV", "Số Hợp Đồng", "Lương TB", "Khen Thưởng", "Kỷ Luật"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu từ bảng
        for (int i = 0; i < modelTongQuat.getRowCount(); i++) {
            Row dataRow = sheet.createRow(rowNum++);

            for (int j = 0; j < modelTongQuat.getColumnCount(); j++) {
                Cell cell = dataRow.createCell(j);
                Object value = modelTongQuat.getValueAt(i, j);

                if (value != null) {
                    if (j == 0) { // Cột tên phòng ban
                        cell.setCellValue(value.toString());
                        cell.setCellStyle(dataStyle);
                    } else { // Các cột số
                        try {
                            double numValue = Double.parseDouble(value.toString().replaceAll("[^0-9.]", ""));
                            cell.setCellValue(numValue);
                            cell.setCellStyle(numberStyle);
                        } catch (NumberFormatException e) {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(dataStyle);
                        }
                    }
                }
            }
        }

        // Điều chỉnh độ rộng cột
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            if (i == 0) { // Cột tên phòng ban rộng hơn
                sheet.setColumnWidth(i, 8000);
            } else {
                sheet.setColumnWidth(i, 3000);
            }
        }
    }

// Các phương thức tạo style
    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Font cho tiêu đề
        org.apache.poi.ss.usermodel.Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setFontName("Arial");
        style.setFont(titleFont);

        return style;
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Font cho header
        org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setFontName("Arial");
        style.setFont(headerFont);

        return style;
    }

    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Font cho dữ liệu
        org.apache.poi.ss.usermodel.Font dataFont = workbook.createFont();
        dataFont.setFontHeightInPoints((short) 11);
        dataFont.setFontName("Arial");
        style.setFont(dataFont);

        return style;
    }

    private CellStyle createNumberStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Format số
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("#,##0"));

        // Font cho số
        org.apache.poi.ss.usermodel.Font numberFont = workbook.createFont();
        numberFont.setFontHeightInPoints((short) 11);
        numberFont.setFontName("Arial");
        style.setFont(numberFont);

        return style;
    }

    private void xuatExcelHopDong() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu báo cáo hợp đồng");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        fileChooser.setSelectedFile(new File("BaoCaoHopDong.xlsx"));

        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();

            // Đảm bảo file có đuôi .xlsx
            if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try {
                // Tạo workbook mới
                Workbook workbook = new XSSFWorkbook();

                // Tạo sheet cho báo cáo hợp đồng
                Sheet sheetHopDong = workbook.createSheet("Báo Cáo Hợp Đồng");
                createHopDongSheet(workbook, sheetHopDong);

                // Tạo sheet thống kê hợp đồng theo loại
                Sheet sheetThongKe = workbook.createSheet("Thống Kê Theo Loại");
                createThongKeHopDongSheet(workbook, sheetThongKe);

                // Ghi file
                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                    workbook.write(outputStream);
                }

                workbook.close();

                JOptionPane.showMessageDialog(this,
                        "Xuất Excel thành công!\nFile đã được lưu tại: " + file.getAbsolutePath(),
                        "Thành công", JOptionPane.INFORMATION_MESSAGE);

            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this,
                        "Lỗi khi xuất Excel: " + e.getMessage(),
                        "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void createHopDongSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style (sử dụng lại các phương thức đã có)
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle dateStyle = createDateStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề chính
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("BÁO CÁO HỢP ĐỒNG LAO ĐỘNG");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

        // Ngày tạo báo cáo và bộ lọc hiện tại
        rowNum++;
        Row dateRow = sheet.createRow(rowNum++);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Ngày tạo báo cáo: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss")));
        dateCell.setCellStyle(dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 6));

        // Thông tin bộ lọc
        Row filterRow = sheet.createRow(rowNum++);
        Cell filterCell = filterRow.createCell(0);
        String filterInfo = "Bộ lọc: " + (String) cmbLoaiHD.getSelectedItem();
        filterCell.setCellValue(filterInfo);
        filterCell.setCellStyle(dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 6));

        rowNum++;

        // Header của bảng
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Mã HĐ", "Tên Nhân Viên", "Loại HĐ", "Ngày Ký", "Ngày Hết Hạn", "Phòng Ban", "Lương CB"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu từ bảng hợp đồng
        for (int i = 0; i < modelHopDong.getRowCount(); i++) {
            Row dataRow = sheet.createRow(rowNum++);

            for (int j = 0; j < modelHopDong.getColumnCount(); j++) {
                Cell cell = dataRow.createCell(j);
                Object value = modelHopDong.getValueAt(i, j);

                if (value != null) {
                    if (j == 3 || j == 4) { // Cột ngày ký và ngày hết hạn
                        if (value instanceof java.util.Date) {
                            cell.setCellValue((java.util.Date) value);
                            cell.setCellStyle(dateStyle);
                        } else {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(dataStyle);
                        }
                    } else if (j == 6) { // Cột lương
                        try {
                            double numValue = Double.parseDouble(value.toString().replaceAll("[^0-9.]", ""));
                            cell.setCellValue(numValue);
                            cell.setCellStyle(numberStyle);
                        } catch (NumberFormatException e) {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(dataStyle);
                        }
                    } else { // Các cột text
                        cell.setCellValue(value.toString());
                        cell.setCellStyle(dataStyle);
                    }
                }
            }
        }

        // Thêm dòng tổng kết
        rowNum++;
        Row summaryRow = sheet.createRow(rowNum++);
        Cell summaryCell = summaryRow.createCell(0);
        summaryCell.setCellValue("Tổng số hợp đồng: " + modelHopDong.getRowCount());
        summaryCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 6));

        // Điều chỉnh độ rộng cột
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            if (i == 1 || i == 2 || i == 5) { // Tên NV, Loại HĐ, Phòng ban
                sheet.setColumnWidth(i, 6000);
            } else if (i == 3 || i == 4) { // Ngày ký, ngày hết hạn
                sheet.setColumnWidth(i, 3500);
            } else {
                sheet.setColumnWidth(i, 3000);
            }
        }
    }

    private void createThongKeHopDongSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);
        CellStyle percentStyle = createPercentStyle(workbook);

        int rowNum = 0;

        // Tiêu đề
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("THỐNG KÊ HỢP ĐỒNG THEO LOẠI");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

        rowNum++;

        // Lấy dữ liệu thống kê
        Map<String, Integer> contractStats = getContractStatistics();
        int totalContracts = contractStats.values().stream().mapToInt(Integer::intValue).sum();

        // Header của bảng thống kê
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Loại Hợp Đồng", "Số Lượng", "Tỷ Lệ (%)", "Ghi Chú"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu thống kê
        for (Map.Entry<String, Integer> entry : contractStats.entrySet()) {
            Row dataRow = sheet.createRow(rowNum++);

            // Loại hợp đồng
            Cell typeCell = dataRow.createCell(0);
            typeCell.setCellValue(entry.getKey());
            typeCell.setCellStyle(dataStyle);

            // Số lượng
            Cell countCell = dataRow.createCell(1);
            countCell.setCellValue(entry.getValue());
            countCell.setCellStyle(numberStyle);

            // Tỷ lệ %
            Cell percentCell = dataRow.createCell(2);
            if (totalContracts > 0) {
                double percent = (double) entry.getValue() / totalContracts;
                percentCell.setCellValue(percent);
                percentCell.setCellStyle(percentStyle);
            } else {
                percentCell.setCellValue(0);
                percentCell.setCellStyle(percentStyle);
            }

            // Ghi chú
            Cell noteCell = dataRow.createCell(3);
            String note = getContractNote(entry.getKey());
            noteCell.setCellValue(note);
            noteCell.setCellStyle(dataStyle);
        }

        // Dòng tổng
        rowNum++;
        Row totalRow = sheet.createRow(rowNum++);
        Cell totalLabelCell = totalRow.createCell(0);
        totalLabelCell.setCellValue("TỔNG CỘNG");
        totalLabelCell.setCellStyle(headerStyle);

        Cell totalCountCell = totalRow.createCell(1);
        totalCountCell.setCellValue(totalContracts);
        totalCountCell.setCellStyle(headerStyle);

        Cell totalPercentCell = totalRow.createCell(2);
        totalPercentCell.setCellValue(1.0);
        totalPercentCell.setCellStyle(headerStyle);

        // Điều chỉnh độ rộng cột
        sheet.setColumnWidth(0, 8000); // Loại hợp đồng
        sheet.setColumnWidth(1, 3000); // Số lượng
        sheet.setColumnWidth(2, 3000); // Tỷ lệ
        sheet.setColumnWidth(3, 6000); // Ghi chú
    }

    private Map<String, Integer> getContractStatistics() {
        Map<String, Integer> stats = new HashMap<>();

        String sql = "SELECT LoaiHD, COUNT(*) AS SoLuong "
                + "FROM Tb_NhanVien "
                + "WHERE LoaiHD IS NOT NULL "
                + "GROUP BY LoaiHD "
                + "ORDER BY SoLuong DESC";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                stats.put(rs.getString("LoaiHD"), rs.getInt("SoLuong"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return stats;
    }

    private String getContractNote(String contractType) {
        switch (contractType) {
            case "Hợp đồng không xác định thời hạn":
                return "Hợp đồng dài hạn";
            case "1 năm":
            case "2 năm":
            case "3 năm":
            case "4 năm":
            case "5 năm":
                return "Hợp đồng xác định thời hạn";
            case "3 tháng":
            case "6 tháng":
                return "Hợp đồng ngắn hạn";
            case "Hợp đồng thử việc":
                return "Thời gian thử việc";
            case "Hợp đồng làm việc bán thời gian":
                return "Làm việc bán thời gian";
            case "Hợp đồng cộng tác viên":
                return "Cộng tác viên";
            default:
                return "Khác";
        }
    }

// Chỉ thêm các phương thức style mới (không trùng lặp)
    private CellStyle createDateStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Format ngày
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("dd/mm/yyyy"));

        // Font
        org.apache.poi.ss.usermodel.Font dateFont = workbook.createFont();
        dateFont.setFontHeightInPoints((short) 11);
        dateFont.setFontName("Arial");
        style.setFont(dateFont);

        return style;
    }

    private CellStyle createPercentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Format phần trăm
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("0.00%"));

        // Font
        org.apache.poi.ss.usermodel.Font percentFont = workbook.createFont();
        percentFont.setFontHeightInPoints((short) 11);
        percentFont.setFontName("Arial");
        style.setFont(percentFont);

        return style;
    }

    private void xuatExcelPhongBan() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Lưu báo cáo phòng ban");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        fileChooser.setSelectedFile(new File("BaoCaoPhongBan.xlsx"));

        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();

            // Đảm bảo file có đuôi .xlsx
            if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try {
                // Tạo workbook mới
                Workbook workbook = new XSSFWorkbook();

                // Tạo sheet cho báo cáo nhân viên theo phòng ban
                Sheet sheetNhanVien = workbook.createSheet("Danh Sách Nhân Viên");
                createNhanVienPhongBanSheet(workbook, sheetNhanVien);

                // Tạo sheet thống kê tổng quan phòng ban
                Sheet sheetThongKe = workbook.createSheet("Thống Kê Phòng Ban");
                createThongKePhongBanSheet(workbook, sheetThongKe);

                // Tạo sheet phân tích lương theo phòng ban
                Sheet sheetLuong = workbook.createSheet("Phân Tích Lương");
                createPhanTichLuongSheet(workbook, sheetLuong);

                // Ghi file
                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                    workbook.write(outputStream);
                }

                workbook.close();

                JOptionPane.showMessageDialog(this,
                        "Xuất Excel thành công!\nFile đã được lưu tại: " + file.getAbsolutePath(),
                        "Thành công", JOptionPane.INFORMATION_MESSAGE);

            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this,
                        "Lỗi khi xuất Excel: " + e.getMessage(),
                        "Lỗi", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void createNhanVienPhongBanSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle dateStyle = createDateStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề chính
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("BÁO CÁO NHÂN VIÊN THEO PHÒNG BAN");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        // Ngày tạo báo cáo và bộ lọc
        rowNum++;
        Row dateRow = sheet.createRow(rowNum++);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Ngày tạo báo cáo: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss")));
        dateCell.setCellStyle(dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

        // Thông tin bộ lọc
        Row filterRow = sheet.createRow(rowNum++);
        Cell filterCell = filterRow.createCell(0);
        String selectedDept = (String) cmbPhongBan.getSelectedItem();
        String filterInfo = "Phòng ban: " + (selectedDept != null ? selectedDept : "Tất cả");
        filterCell.setCellValue(filterInfo);
        filterCell.setCellStyle(dataStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

        rowNum++;

        // Header của bảng
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Mã NV", "Tên Nhân Viên", "Chức Vụ", "Ngày Vào Làm", "Lương CB", "Loại HĐ"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu từ bảng nhân viên phòng ban
        for (int i = 0; i < modelPhongBan.getRowCount(); i++) {
            Row dataRow = sheet.createRow(rowNum++);

            for (int j = 0; j < modelPhongBan.getColumnCount(); j++) {
                Cell cell = dataRow.createCell(j);
                Object value = modelPhongBan.getValueAt(i, j);

                if (value != null) {
                    if (j == 3) { // Cột ngày vào làm
                        if (value instanceof java.util.Date) {
                            cell.setCellValue((java.util.Date) value);
                            cell.setCellStyle(dateStyle);
                        } else {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(dataStyle);
                        }
                    } else if (j == 4) { // Cột lương CB
                        try {
                            double numValue = Double.parseDouble(value.toString().replaceAll("[^0-9.]", ""));
                            cell.setCellValue(numValue);
                            cell.setCellStyle(numberStyle);
                        } catch (NumberFormatException e) {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(dataStyle);
                        }
                    } else { // Các cột text
                        cell.setCellValue(value.toString());
                        cell.setCellStyle(dataStyle);
                    }
                }
            }
        }

        // Thêm dòng tổng kết
        rowNum++;
        Row summaryRow = sheet.createRow(rowNum++);
        Cell summaryCell = summaryRow.createCell(0);
        summaryCell.setCellValue("Tổng số nhân viên: " + modelPhongBan.getRowCount());
        summaryCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

        // Điều chỉnh độ rộng cột
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            if (i == 1) { // Tên nhân viên
                sheet.setColumnWidth(i, 6000);
            } else if (i == 2 || i == 5) { // Chức vụ, Loại HĐ
                sheet.setColumnWidth(i, 5000);
            } else if (i == 3) { // Ngày vào làm
                sheet.setColumnWidth(i, 3500);
            } else {
                sheet.setColumnWidth(i, 3000);
            }
        }
    }

    private void createThongKePhongBanSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("THỐNG KÊ TỔNG QUAN THEO PHÒNG BAN");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        rowNum++;

        // Lấy dữ liệu thống kê phòng ban
        Map<String, DepartmentStats> deptStats = getDepartmentStatistics();

        // Header của bảng thống kê
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Phòng Ban", "Số NV", "Số HĐ", "Lương TB", "Khen Thưởng", "Kỷ Luật"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu thống kê
        int totalEmployees = 0;
        int totalContracts = 0;
        double totalSalary = 0;
        int totalRewards = 0;
        int totalPunishments = 0;

        for (Map.Entry<String, DepartmentStats> entry : deptStats.entrySet()) {
            Row dataRow = sheet.createRow(rowNum++);
            DepartmentStats stats = entry.getValue();

            // Tên phòng ban
            Cell deptCell = dataRow.createCell(0);
            deptCell.setCellValue(entry.getKey());
            deptCell.setCellStyle(dataStyle);

            // Số nhân viên
            Cell empCell = dataRow.createCell(1);
            empCell.setCellValue(stats.employeeCount);
            empCell.setCellStyle(numberStyle);

            // Số hợp đồng
            Cell contractCell = dataRow.createCell(2);
            contractCell.setCellValue(stats.contractCount);
            contractCell.setCellStyle(numberStyle);

            // Lương trung bình
            Cell salaryCell = dataRow.createCell(3);
            salaryCell.setCellValue(stats.averageSalary);
            salaryCell.setCellStyle(numberStyle);

            // Khen thưởng
            Cell rewardCell = dataRow.createCell(4);
            rewardCell.setCellValue(stats.rewardCount);
            rewardCell.setCellStyle(numberStyle);

            // Kỷ luật
            Cell punishCell = dataRow.createCell(5);
            punishCell.setCellValue(stats.punishmentCount);
            punishCell.setCellStyle(numberStyle);

            // Cộng dồn tổng
            totalEmployees += stats.employeeCount;
            totalContracts += stats.contractCount;
            totalSalary += stats.averageSalary * stats.employeeCount;
            totalRewards += stats.rewardCount;
            totalPunishments += stats.punishmentCount;
        }

        // Dòng tổng
        rowNum++;
        Row totalRow = sheet.createRow(rowNum++);

        Cell totalLabelCell = totalRow.createCell(0);
        totalLabelCell.setCellValue("TỔNG CỘNG");
        totalLabelCell.setCellStyle(headerStyle);

        Cell totalEmpCell = totalRow.createCell(1);
        totalEmpCell.setCellValue(totalEmployees);
        totalEmpCell.setCellStyle(headerStyle);

        Cell totalContractCell = totalRow.createCell(2);
        totalContractCell.setCellValue(totalContracts);
        totalContractCell.setCellStyle(headerStyle);

        Cell avgSalaryCell = totalRow.createCell(3);
        double overallAvgSalary = totalEmployees > 0 ? totalSalary / totalEmployees : 0;
        avgSalaryCell.setCellValue(overallAvgSalary);
        avgSalaryCell.setCellStyle(headerStyle);

        Cell totalRewardCell = totalRow.createCell(4);
        totalRewardCell.setCellValue(totalRewards);
        totalRewardCell.setCellStyle(headerStyle);

        Cell totalPunishCell = totalRow.createCell(5);
        totalPunishCell.setCellValue(totalPunishments);
        totalPunishCell.setCellStyle(headerStyle);

        // Điều chỉnh độ rộng cột
        sheet.setColumnWidth(0, 8000); // Phòng ban
        for (int i = 1; i < headers.length; i++) {
            sheet.setColumnWidth(i, 3000);
        }
    }

    private void createPhanTichLuongSheet(Workbook workbook, Sheet sheet) {
        // Tạo các style
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);

        int rowNum = 0;

        // Tiêu đề
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("PHÂN TÍCH LƯƠNG THEO PHÒNG BAN");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        rowNum++;

        // Lấy dữ liệu phân tích lương
        Map<String, SalaryAnalysis> salaryData = getSalaryAnalysis();

        // Header của bảng
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"Phòng Ban", "Lương Thấp Nhất", "Lương Cao Nhất", "Lương Trung Bình", "Độ Lệch Chuẩn"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Dữ liệu phân tích
        for (Map.Entry<String, SalaryAnalysis> entry : salaryData.entrySet()) {
            Row dataRow = sheet.createRow(rowNum++);
            SalaryAnalysis analysis = entry.getValue();

            // Tên phòng ban
            Cell deptCell = dataRow.createCell(0);
            deptCell.setCellValue(entry.getKey());
            deptCell.setCellStyle(dataStyle);

            // Lương thấp nhất
            Cell minCell = dataRow.createCell(1);
            minCell.setCellValue(analysis.minSalary);
            minCell.setCellStyle(numberStyle);

            // Lương cao nhất
            Cell maxCell = dataRow.createCell(2);
            maxCell.setCellValue(analysis.maxSalary);
            maxCell.setCellStyle(numberStyle);

            // Lương trung bình
            Cell avgCell = dataRow.createCell(3);
            avgCell.setCellValue(analysis.avgSalary);
            avgCell.setCellStyle(numberStyle);

            // Độ lệch chuẩn
            Cell stdCell = dataRow.createCell(4);
            stdCell.setCellValue(analysis.stdDeviation);
            stdCell.setCellStyle(numberStyle);
        }

        // Điều chỉnh độ rộng cột
        sheet.setColumnWidth(0, 8000); // Phòng ban
        for (int i = 1; i < headers.length; i++) {
            sheet.setColumnWidth(i, 4000);
        }
    }

// Lớp để lưu thống kê phòng ban
    private static class DepartmentStats {

        int employeeCount;
        int contractCount;
        double averageSalary;
        int rewardCount;
        int punishmentCount;

        DepartmentStats(int empCount, int contractCount, double avgSalary, int rewards, int punishments) {
            this.employeeCount = empCount;
            this.contractCount = contractCount;
            this.averageSalary = avgSalary;
            this.rewardCount = rewards;
            this.punishmentCount = punishments;
        }
    }

// Lớp để lưu phân tích lương
    private static class SalaryAnalysis {

        double minSalary;
        double maxSalary;
        double avgSalary;
        double stdDeviation;

        SalaryAnalysis(double min, double max, double avg, double std) {
            this.minSalary = min;
            this.maxSalary = max;
            this.avgSalary = avg;
            this.stdDeviation = std;
        }
    }

    private Map<String, DepartmentStats> getDepartmentStatistics() {
        Map<String, DepartmentStats> stats = new HashMap<>();

        String sql = "SELECT pb.TenPB, "
                + "COUNT(DISTINCT nv.MaNhanVien) AS SoNV, "
                + "COUNT(DISTINCT hd.MaHD) AS SoHD, "
                + "COALESCE(AVG(bl.LuongCoBan), 0) AS LuongTB, "
                + "COUNT(DISTINCT kt.MaKT) AS SoKT, "
                + "COUNT(DISTINCT kl.MaKL) AS SoKL "
                + "FROM Tb_PhongBan pb "
                + "LEFT JOIN Tb_NhanVien nv ON pb.TenPB = nv.PhongBan "
                + "LEFT JOIN Tb_HopDong hd ON nv.MaNhanVien = hd.MaNhanVien "
                + "LEFT JOIN Tb_BangLuong bl ON nv.MaNhanVien = bl.MaNhanVien "
                + "LEFT JOIN Tb_KhenThuong kt ON nv.MaNhanVien = kt.MaNhanVien "
                + "LEFT JOIN Tb_KyLuat kl ON nv.MaNhanVien = kl.MaNhanVien "
                + "GROUP BY pb.MaPB, pb.TenPB "
                + "ORDER BY pb.TenPB";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                String deptName = rs.getString("TenPB");
                DepartmentStats deptStats = new DepartmentStats(
                        rs.getInt("SoNV"),
                        rs.getInt("SoHD"),
                        rs.getDouble("LuongTB"),
                        rs.getInt("SoKT"),
                        rs.getInt("SoKL")
                );
                stats.put(deptName, deptStats);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return stats;
    }

    private Map<String, SalaryAnalysis> getSalaryAnalysis() {
        Map<String, SalaryAnalysis> analysis = new HashMap<>();

        String sql = "SELECT nv.PhongBan, "
                + "MIN(bl.LuongCoBan) AS LuongMin, "
                + "MAX(bl.LuongCoBan) AS LuongMax, "
                + "AVG(bl.LuongCoBan) AS LuongTB, "
                + "STDEV(bl.LuongCoBan) AS DoCech "
                + "FROM Tb_NhanVien nv "
                + "INNER JOIN Tb_BangLuong bl ON nv.MaNhanVien = bl.MaNhanVien "
                + "WHERE bl.LuongCoBan IS NOT NULL "
                + "GROUP BY nv.PhongBan "
                + "ORDER BY nv.PhongBan";

        try (Connection conn = connection.getConnection(); PreparedStatement pstmt = conn.prepareStatement(sql); ResultSet rs = pstmt.executeQuery()) {

            while (rs.next()) {
                String deptName = rs.getString("PhongBan");
                SalaryAnalysis salaryAnalysis = new SalaryAnalysis(
                        rs.getDouble("LuongMin"),
                        rs.getDouble("LuongMax"),
                        rs.getDouble("LuongTB"),
                        rs.getDouble("DoCech")
                );
                analysis.put(deptName, salaryAnalysis);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return analysis;
    }

    private void taoBaoCaoTongHop() {
        JOptionPane.showMessageDialog(this,
                "Tính năng tạo báo cáo tổng hợp đang được phát triển!\n"
                + "Sẽ bao gồm:\n"
                + "- Báo cáo chi tiết toàn bộ nhân sự\n"
                + "- Phân tích xu hướng tuyển dụng\n"
                + "- Đánh giá hiệu suất theo phòng ban\n"
                + "- Dự báo nhân sự",
                "Thông báo", JOptionPane.INFORMATION_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new BaoCaoThongKe().setVisible(true);
        });
    }
}
