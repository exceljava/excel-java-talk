/**
 Trivial example without using vol surfaces or swap curves.
 This is not to be meant as 'real' code!

 If you are interested in a real example using financial models
 take a look at https://github.com/exceljava/strata-excel.
 */

package com.exceljava.example;

import java.util.HashMap;

public class MarketDataMock implements MarketData {

    private HashMap<String, Double> spotPrices = new HashMap<>();
    private HashMap<String, Double> volatilities = new HashMap<>();
    private HashMap<String, Double> dividendYields = new HashMap<>();
    private HashMap<String, String> currencies = new HashMap<>();
    private HashMap<String, Double> interestRates = new HashMap<>();

    public MarketDataMock() {
    }

    public void addTicker(String ticker, String currency, double spotPrice, double volatilty, double dividendYield) {
        spotPrices.put(ticker, spotPrice);
        volatilities.put(ticker, volatilty);
        dividendYields.put(ticker, dividendYield);
        currencies.put(ticker, currency);
    }

    public void addInterestRate(String currency, double interestRate) {
        interestRates.put(currency, interestRate);
    }

    @Override
    public double getSpot(String ticker) {
        return spotPrices.get(ticker);
    }

    @Override
    public double getVolatility(String ticker, double strike, double yearsToExpiry) {
        return volatilities.get(ticker);
    }

    @Override
    public String getCurrency(String ticker) {
        return currencies.get(ticker);
    }

    @Override
    public double getInterestRate(String currency, double yearsToExpiry) {
        return interestRates.get(currency);
    }

    @Override
    public double getDividendYield(String ticker) {
        return dividendYields.get(ticker);
    }
}
