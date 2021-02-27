package com.romel.cathlabcalloff;

public class CathLabController {

    private CathLabView view;

    public CathLabController() {
        view = new CathLabView();
    }

    public void run() {
        initListeners();
        view.showGui();
    }

    private void initListeners() {
        view.getButton().addActionListener((e) -> {
            view.getLabel().setText(">>Status: Running.");
            CathLabService service = new CathLabService();
            service.run();
            view.getLabel().setText(">>Status: Completed.");
        });
    }

}