package com.exceljava;

import com.opencsv.bean.CsvBindByName;

public class Record {
    @CsvBindByName(column="Portfolio")
    private String portfolio;

    @CsvBindByName(column="Security")
    private String security;

    @CsvBindByName(column="Notional")
    private double notional;

    @CsvBindByName(column="Market Value")
    private double marketValue;

    @CsvBindByName(column="Delta %")
    private double deltaPct;

    @CsvBindByName(column="Delta $")
    private double deltaValue;

    @CsvBindByName(column="1Y")
    private double riskFactor1y;

    @CsvBindByName(column="3Y")
    private double riskFactor3y;

    @CsvBindByName(column="5Y")
    private double riskFactor5y;

    @CsvBindByName(column="10Y")
    private double riskFactor10y;

    @CsvBindByName(column="15Y")
    private double riskFactor15y;

    public String getPortfolio() {
        return portfolio;
    }

    public String getSecurity() {
        return security;
    }

    public double getNotional() {
        return notional;
    }

    public double getMarketValue() {
        return marketValue;
    }

    public double getDeltaPct() {
        return deltaPct;
    }

    public double getDeltaValue() {
        return deltaValue;
    }

    public double getRiskFactor1y() {
        return riskFactor1y;
    }

    public double getRiskFactor3y() {
        return riskFactor3y;
    }

    public double getRiskFactor5y() {
        return riskFactor5y;
    }

    public double getRiskFactor10y() {
        return riskFactor10y;
    }

    public double getRiskFactor15y() {
        return riskFactor15y;
    }
}
