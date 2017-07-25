
makePackedBars = function(inArgs) {
    var arg = inArgs;
    if (arg.csvFile !== null) {
        d3.csv(arg.csvFile, function (d) {
                return d;
            },
            function (error, rawdata) {
                if (error) throw error;
                arg.rawData = rawdata;
                makePackedBarsFromData(arg);
            }
        );
    }
    else {
        makePackedBarsFromData(arg);
    }
}

makePackedBarsFromData = function(inArgs) {
    // default parameters: null values will be computed if not supplied
    var arg = {
        nPrimary: 15,
        categoryDataName: null,
        valueDataName: null,
        //categoryTitle: null,
        valueTitle: null,
        chartTitle: null,
        rawData: null,
        csvFile: null,
        chartMargin: {top: 30, right: 20, bottom: 50, left: 20},
        //chartWidth: 900,
        //chartHeight: 500,
        labelThreshold: 0.7 // threshold for labeling secondary bars
    };
    for (var propName in inArgs) {
        arg[propName] = inArgs[propName];
    }

    if (arg.categoryDataName === null)
        arg.categoryDataName = arg.rawData.columns[0];
    if (arg.valueDataName === null)
        arg.valueDataName = arg.rawData.columns[1];
    if (arg.categoryTitle === null)
        arg.categoryTitle = arg.categoryDataName;
    if (arg.valueTitle === null)
        arg.valueTitle = arg.valueDataName;

    var data = arg.rawData.map(function (rd) {
        return {
            category: rd[arg.categoryDataName],
            value: +rd[arg.valueDataName]
        };
    });
    data.sort(function (a, b) {
        return b.value - a.value
    });  // descending order

    var nCategory = data.length;
    var nRows = Math.min(nCategory, arg.nPrimary);
    if (arg.chartTitle === null)
        arg.chartTitle = "Top " + nRows + " of " + nCategory + " " + arg.categoryDataName + " " + arg.valueDataName + " values";

    var minBarHeight = 15;
    var minHeight = minBarHeight * nRows;
    var index = Array.from({length: nCategory}, function (v, k) {
        return k;
    });
    var minPrimary = data[nRows - 1].value;

    var svg = d3.select("svg"),
        margin = arg.chartMargin,
        width = +svg.attr("width") - margin.left - margin.right,
        height = d3.max([+svg.attr("height"), minHeight]) - margin.top - margin.bottom;

    svg.style("height", height + margin.top + margin.bottom);

    var g = svg.append("g")
        .attr("transform", "translate(" + margin.left + "," + margin.top + ")");
    var x = d3.scaleLinear().rangeRound([0, width]),
        y = d3.scaleBand().range([0, height]).paddingOuter(0).paddingInner((nRows + 1) / height);

    x.domain([0, d3.max([data[0].value, d3.sum(data, function (d) {
        return d.value;
    }) / nRows])]);
    //y.domain([0, nRows-1]);    //data.map(function(d) { return d.category; }));
    y.domain(Array.from({length: nRows}, function (v, k) {
        return k;
    }));

    // debugging scratch
    var dd = data[1].category;
    var ee = y(0);
    var tt = y(1);
    var ww = y.bandwidth();

    var xax = g.append("g")
        .attr("class", "axis axis--x")
        .attr("transform", "translate(0," + height + ")")
        .call(d3.axisBottom(x));

    g.append("text")
        .attr("class", "title x-title")
        .attr("x", width / 2)
        .attr("y", height + xax.node().getBBox().height + 6)
        .attr("text-anchor", "middle")
        .attr("alignment-baseline", "hanging")
        .text(arg.valueTitle);

    g.append("text")
        .attr("class", "title graph-title")
        .attr("x", width / 2)
        .attr("y", -10)
        .attr("text-anchor", "middle")
        .text(arg.chartTitle);

    var rowSum = new Array(nRows);  // running for each row
    var rowGray = new Array(nRows); // last used gray for each row, to prevent dups
    var ic;
    for (ic = 0; ic < nRows; ic++) {
        data[ic].xlo = 0;
        data[ic].row = ic;
        rowSum[ic] = data[ic].value;
        data[ic].gray = 0;
        rowGray[ic] = -1;
    }
    for (ic = nRows; ic < nCategory; ic++) {
        var ir = d3.scan(rowSum);
        data[ic].xlo = rowSum[ir];
        data[ic].row = ir;
        rowSum[ir] += data[ic].value;
        data[ic].gray = ic % 4;
        if (data[ic].gray == rowGray[ir])
            data[ic].gray = (data[ic].gray + 1) % 4;
        rowGray[ir] = data[ic].gray;
    }

    var enterSelection = g.selectAll(".bar").data(index).enter();
    enterSelection.append("rect")
        .attr("class", function (i) {
            return i < nRows ? "bar1" : "bar2";
        })
        .attr("fill", function (i) {
            var g = 237 - 4 * data[i].gray;   // a few light grays
            return i < nRows ? "#5b76c0" : "rgb(" + g + "," + g + ", " + g + ")";
        })
        .attr("x", function (i) {
            return x(data[i].xlo);
        })
        .attr("y", function (i) {
            return y(data[i].row);
        })
        .attr("width", function (i) {
            return x(data[i].value);
        })
        .attr("height", y.bandwidth())
        .append("title")
        .text(function (i) {
            return data[i].category;
        });
    enterSelection.insert("text")
        .attr("class", function (i) {
            return i < nRows ? "label1" : "label2";
        })
        .attr("fill", function (i) {
            return i < nRows ? "#fff" : "#888";
        })
        .attr("x", function (i) {
            return i < nRows ? x(data[i].xlo) + 5 : x(data[i].xlo) + x(data[i].value) / 2;
        })
        .attr("y", function (i) {
            return y(data[i].row) + y.bandwidth() / 2 + 5;
        })
        .attr("alignment-baseline", "middle")
        .attr("text-anchor", function (i) {
            return i < nRows ? "start" : "middle";
        })
        .append('tspan')
        .text(function (i) {
            return data[i].value >= arg.labelThreshold * minPrimary ? data[i].category : "";
        })
        .attr('width', function (i) {
            return x(data[i].value);
        })
        .each(wrap);

    function wrap(d) {
        var self = d3.select(this),
            textLength = self.node().getComputedTextLength(),
            text = self.text();
        while (( textLength > self.attr('width') ) && text.length > 0) {
            text = text.slice(0, -1);
            self.text(text + '...');
            textLength = self.node().getComputedTextLength();
        }
    }

}