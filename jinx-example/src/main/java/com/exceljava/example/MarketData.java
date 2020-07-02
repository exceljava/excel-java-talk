/**
 Trivial example without using vol surfaces or swap curves.
 This is not to be meant as 'real' code!

 If you are interested in a real example using financial models
 take a look at https://github.com/exceljava/strata-excel.
 */

package com.exceljava.example;

public interface MarketData {
    /**
     * Get the spot price for a stock.
     * @param ticker Ticket to get spot price of.
     * @return Spot price.
     */
    public double getSpot(String ticker);

    /**
     * Get volatility for using with Black-Scholes option valuation.
     * @param ticker Stock ticker
     * @param strike Option strike price
     * @param yearsToExpiry Years to option expiry
     * @return Implied volatility
     */
    public double getVolatility(String ticker, double strike, double yearsToExpiry);

    /**
     * Get the currency from a stock ticker.
     * @param ticker Ticket to get currency of.
     * @return Currency as a string.
     */
    public String getCurrency(String ticker);

    /**
     * Get the simple interest rate for a currency for option valuation.
     * @param currency Name of currency.
     * @param yearsToExpiry Years to option expiry.
     * @return Interest rate.
     */
    public double getInterestRate(String currency, double yearsToExpiry);

    /**
     * Get the simple dividend yield for a stock to use for option valuation.
     * @param ticker Stock ticker
     * @return Dividend yield.
     */
    public double getDividendYield(String ticker);
}
