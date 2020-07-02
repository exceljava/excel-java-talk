package com.exceljava.example;

import com.exceljava.jinx.ExcelFunction;
import org.apache.commons.math3.distribution.NormalDistribution;

public class BlackScholes {
    @ExcelFunction
    public static double blackScholes(double spot,
                               double strike,
                               double yearsToExpiry,
                               double volatility,
                               double interestRate,
                               double dividendYield,
                               boolean isCall) {
        if (yearsToExpiry <= 0.0) {
            throw new IllegalArgumentException("Years to expiry must be > 0");
        }

        if (volatility <= 0.0) {
            throw new IllegalArgumentException("Volatility must be > 0");
        }

        NormalDistribution dist = new NormalDistribution();
        double d1 = (Math.log(spot / strike) + (interestRate - dividendYield + 0.5 * Math.pow(volatility, 2)) * yearsToExpiry) / (volatility * Math.sqrt(yearsToExpiry));
        double d2 = (Math.log(spot / strike) + (interestRate - dividendYield - 0.5 * Math.pow(volatility, 2)) * yearsToExpiry) / (volatility * Math.sqrt(yearsToExpiry));

        if (isCall) {
            return spot * Math.exp(-dividendYield * yearsToExpiry) * dist.cumulativeProbability(d1)
                    - strike * Math.exp(-interestRate * yearsToExpiry) * dist.cumulativeProbability(d2);
        }

        return strike * Math.exp(-interestRate * yearsToExpiry) * dist.cumulativeProbability(-d2)
                - spot * Math.exp(-dividendYield * yearsToExpiry) * dist.cumulativeProbability(-d1);
    }

    @ExcelFunction
    public static MarketData loadMarketData() {
        MarketDataMock marketData = new MarketDataMock();
        marketData.addTicker("XYZ", "USD", 250, 0.2, 0.01);
        marketData.addInterestRate("USD", 0.01);
        return marketData;
    }

    @ExcelFunction
    public static double blackScholes2(MarketData marketData,
                                       String ticker,
                                       double strike,
                                       double yearsToExpiry,
                                       boolean isCall) {
        double spot = marketData.getSpot(ticker);
        double volatility = marketData.getVolatility(ticker, strike, yearsToExpiry);
        String currency = marketData.getCurrency(ticker);
        double interestRate = marketData.getInterestRate(currency, yearsToExpiry);
        double dividendYield = marketData.getDividendYield(ticker);

        return blackScholes(spot,
                strike,
                yearsToExpiry,
                volatility,
                interestRate,
                dividendYield,
                isCall);
    }
}
