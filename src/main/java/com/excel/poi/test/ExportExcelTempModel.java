package com.excel.poi.test;

import java.util.Date;

public class ExportExcelTempModel {

    private String locationName;

    private String deptName;

    private String money;

    private String number;

    private Date time;

    public ExportExcelTempModel() {
    }

    public ExportExcelTempModel(String locationName, String deptName, String money, String number, Date time) {
        this.locationName = locationName;
        this.deptName = deptName;
        this.money = money;
        this.number = number;
        this.time = time;
    }

    public String getLocationName() {
        return locationName;
    }

    public void setLocationName(String locationName) {
        this.locationName = locationName;
    }

    public String getDeptName() {
        return deptName;
    }

    public void setDeptName(String deptName) {
        this.deptName = deptName;
    }

    public String getMoney() {
        return money;
    }

    public void setMoney(String money) {
        this.money = money;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public Date getTime() {
        return time;
    }

    public void setTime(Date time) {
        this.time = time;
    }
}
