package com.romel.cathlabcalloff;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;

public class CathLabView {

    private JFrame frame;
    private JPanel panel;
    private JPanel panelButton;
    private JButton button;
    private JLabel label;

    public CathLabView() {
        init();
    }

    public JButton getButton() {
        return button;
    }

    public JFrame getFrame() {
        return frame;
    }

    public JLabel getLabel() {
        return label;
    }

    public void init() {
        frame = new JFrame("CathLab Call Off Report");
        frame.setSize(300, 100);
        frame.setLayout(new BorderLayout(5, 5));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLocationRelativeTo(null);

        button = new JButton("Start");
        label = new JLabel(">Status:");

        panel = new JPanel(new BorderLayout(5, 5));
        frame.getContentPane().add(panel, BorderLayout.CENTER);

        panelButton = new JPanel(new FlowLayout());
        panelButton.add(button);
        panel.add(panelButton, BorderLayout.CENTER);
        panel.add(label, BorderLayout.PAGE_END);
    }

    public void showGui() {
        frame.setVisible(true);
    }

}