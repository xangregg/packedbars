## D3 implementation of Packed Bars chart

Simple usage:

```javascript
<svg width="960" height="500"></svg>
<script src="https://d3js.org/d3.v4.min.js"></script>
<script src="makepackedbars.js"></script>
<script>
    makePackedBars({ csvFile: "sp500.csv" });
</script>

```

Fancier usage:

```javascript
<svg width="960" height="500"></svg>
<script src="https://d3js.org/d3.v4.min.js"></script>
<script src="makepackedbars.js"></script>
<script>
    makePackedBars({
        csvFile: "sp500.csv",
        nPrimary: 15,
        valueTitle: "Market Capitalization in $billions",
        chartTitle: "S&P 500 Market Caps",
        chartMargin: {top: 30, right: 20, bottom: 50, left: 20},
        labelThreshold: 0.6 // threshold for labeling secondary bars
    });
</script>

```

