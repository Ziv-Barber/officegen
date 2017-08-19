var officegen = require('../');
var OfficeChart = require('../lib/officechart.js');
var async = require('async');

var fs = require('fs');
var path = require('path');

var pptx = officegen('pptx');

var slide;
var pObj;

pptx.on('finalize', function (written) {
  console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n');

  // clear the temporatory files
});

pptx.on('error', function (err) {
  console.log(err);
});

pptx.setDocTitle('Sample PPTX Document');


// this shows how one can get the base XML and modify it directly
/*
var chart0 = new OfficeChart({
  title: 'Dynamically generated',
  renderType: 'bar',
  overlap: 50,
  gapWidth: 25,
  valAxisTitle: "Da Value Axis",
  catAxisTitle: "Da Cat Axis",
  catAxisReverseOrder: true,
  valAxisCrossAtMaxCategory: true,
  valAxisMajorGridlines: true,
  valAxisMinorGridlines: true,
  data: [
    {
      name: 'Income',
      labels: ['2005', '2006', '2007', '2008', '2009'],
      values: [23.5, 26.2, 30.1, 29.5, 24.6],
      // schemeColor: 'accent1'
      // color: 'ff0000',
      xml: {
        "c:spPr": {
          "a:solidFill": {
            "a:schemeClr": { "@val": "accent1"}
          },
          "a:ln": {
            "a:solidFill": {
              "a:schemeClr": { "@val": "tx1"}
            }
          }
        }
      }
    },
    {
      name: 'Expense',
      labels: ['2005', '2006', '2007', '2008', '2009'],
      values: [18.1, 22.8, 23.9, 25.1, 25],
      // color: '00ff00',
      // schemeColor: 'bg2'
      xml: {
        "c:spPr": {
          "a:solidFill": {
            "a:schemeClr": { "@val": "bg2"}
          },
          "a:ln": {
            "a:solidFill": {
              "a:schemeClr": { "@val": "tx1"}
            }
          }
        }
      }
    }
  ],
  fontSize: "1200", // equivalent to specifying the xml below
  xml: {
      "c:txPr": {
        "a:bodyPr": {},
        "a:listStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {
              "@sz": "1200"
            }
          },
          "a:endParaRPr": {
            "@lang": "en-US"
          }
        }
      }
    }
});
*/

var chartsData = [
  // chart0,
  {
    "title": "Marginal distribution for mpg",
    "renderType": "column",
    "valAxisNumFmt": "0%",
    valAxisMaxValue: 24,
    "data": [
      {
        "name": "current",
        "labels": [
          "[NA]",
          "14.1 to 16",
          "16.1 to 18",
          "18.1 to 20",
          "20.1 to 22",
          "22.1 to 24",
          "24.1 to 26",
          "26.1 to 28",
          "28.1 to 30",
          "30.1 to 32",
          "32.1 to 34",
          "44.1 to 46"
        ],
        "values": [
          0.024390243902439025,
          0.17073170731707318,
          0.1951219512195122,
          0.21951219512195122,
          0.14634146341463414,
          0.21951219512195122,
          0,
          0.024390243902439025,
          0,
          0,
          0,
          0
        ],
        "xml": {
          "c:spPr": {
            "a:solidFill": {
              "a:schemeClr": {
                "@val": "accent1"
              }
            },
            "a:ln": {
              "a:solidFill": {
                "a:schemeClr": {
                  "@val": "tx1"
                }
              }
            }
          }

        }
      },

      {
        "name": "baseline",
        "labels": [
          "[NA]",
          "14.1 to 16",
          "16.1 to 18",
          "18.1 to 20",
          "20.1 to 22",
          "22.1 to 24",
          "24.1 to 26",
          "26.1 to 28",
          "28.1 to 30",
          "30.1 to 32",
          "32.1 to 34",
          "44.1 to 46"
        ],
        "values": [
          0.017241379310344827,
          0.008620689655172414,
          0,
          0.017241379310344827,
          0.1896551724137931,
          0.1810344827586207,
          0.29310344827586204,
          0.15517241379310345,
          0.0603448275862069,
          0.034482758620689655,
          0.034482758620689655,
          0.008620689655172414
        ],
        "xml": {
          "c:spPr": {
            "a:solidFill": {
              "a:schemeClr": {
                "@val": "bg2"
              }
            },
            "a:ln": {
              "a:solidFill": {
                "a:schemeClr": {
                  "@val": "tx1"
                }
              }
            }
          }
        }
      }
    ]
  },

  {
    title: 'My production',
    renderType: 'pie',
    data: [
      {
        name: 'Oil',
        labels: ['Czech Republic', 'Ireland', 'Germany', 'Australia', 'Austria', 'UK', 'Belgium'],
        values: [301, 201, 165, 139, 128, 99, 60],
        colors: ['ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff', '000000']
      }
    ]
  },
  {
    title: 'My production',
    renderType: 'doughnut',
    data: [
      {
        name: 'Oil',
        labels: ['Czech Republic', 'Ireland', 'Germany', 'Australia', 'Austria', 'UK', 'Belgium'],
        values: [301, 201, 165, 139, 128, 99, 60],
        colors: ['ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff', '000000']
      }
    ]
  },
  {
    title: 'line chart',
    renderType: 'line',
    data: [
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.6, 2.8],
        color: 'ff0000'
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.7, 2.9],
        color: '00ff00'
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.1, 2.2, 2.4],
        color: '0000ff'
      },
      {
        name: 'lamerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.3, 0.3, 0.3],
        color: 'ffff00'
      },
      {
        name: 'meast',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.2, 0.3, 0.3],
        color: 'ff00ff'
      },
      {
        name: 'africa',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.1, 0.1, 0.1],
        color: '00ffff'
      }
    ]
  },

  {
    title: 'eSurvey chart',
    renderType: 'column',
    overlap: 50,
    gapWidth: 25,
    valAxisNumFmt: '$0',
    valAxisMaxValue: 24,
    data: [
      {
        name: 'Income',
        labels: ['2005', '2006', '2007', '2008', '2009'],
        values: [23.5, 26.2, 30.1, 29.5, 24.6],
        color: 'ff0000'
      },
      {
        name: 'Expense',
        labels: ['2005', '2006', '2007', '2008', '2009'],
        values: [18.1, 22.8, 23.9, 25.1, 25],
        color: '00ff00'
      }
    ]
  },
  {
    title: 'eSurvey chart',
    renderType: 'line',
    overlap: 50,
    gapWidth: 25,
    valAxisNumFmt: '$0',
    valAxisMaxValue: 24,
    data: [
      {
        name: 'Income',
        labels: ['2005', '2006', '2007', '2008', '2009'],
        values: [23.5, 26.2, 30.1, 29.5, 24.6],
        color: 'ff0000'
      },
      {
        name: 'Expense',
        labels: ['2005', '2006', '2007', '2008', '2009'],
        values: [18.1, 22.8, 23.9, 25.1, 25],
        color: '00ff00'
      }
    ]
  },
  {
    title: 'eSurvey chart',
    renderType: 'stacked-column',
    valAxisNumFmt: '$0',
    valAxisMaxValue: 24,
    data: [
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.6, 2.8],
        color: 'ff0000'
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.7, 2.9],
        color: '00ff00'
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.1, 2.2, 2.4],
        color: '0000ff'
      },
      {
        name: 'lamerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.3, 0.3, 0.3],
        color: 'ffff00'
      },
      {
        name: 'meast',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.2, 0.3, 0.3],
        color: 'ff00ff'
      },
      {
        name: 'africa',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.1, 0.1, 0.1],
        color: '00ffff'
      }

    ]
  },

  {
    title: 'Sample bar chart',
    renderType: 'bar',
    xmlOptions: {
      "c:title": {
        "c:tx": {
          "c:rich": {
            "a:p": {
              "a:r": {
                "a:t": "Override title via XML"
              }
            }
          }
        }
      }
    },
    data: [
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.6, 2.8],
        color: 'ff0000'
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.7, 2.9],
        color: '00ff00'
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.1, 2.2, 2.4],
        color: '0000ff'
      },
      {
        name: 'lamerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.3, 0.3, 0.3],
        color: 'ffff00'
      },
      {
        name: 'meast',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.2, 0.3, 0.3],
        color: 'ff00ff'
      },
      {
        name: 'africa',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.1, 0.1, 0.1],
        color: '00ffff'
      }

    ]
  },

  {
    title: 'Group bar chart',
    renderType: 'group-bar',
    data: [
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.6, 2.8],
        color: 'ff0000'
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.7, 2.9],
        color: '00ff00'
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.1, 2.2, 2.4],
        color: '0000ff'
      },
      {
        name: 'lamerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.3, 0.3, 0.3],
        color: 'ffff00'
      },
      {
        name: 'meast',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.2, 0.3, 0.3],
        color: 'ff00ff'
      },
      {
        name: 'africa',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.1, 0.1, 0.1],
        color: '00ffff'
      }
    ]
  }
];


function generateOneChart(chartInfo, callback) {

  slide = pptx.makeNewSlide();
  slide.name = 'OfficeChart slide';
  slide.back = 'ffffff';
  slide.addChart(chartInfo, callback, callback);
}

function generateCharts(callback) {
  async.each(chartsData, generateOneChart, callback);
}


function finalize() {
  var out = fs.createWriteStream('tmp/out_charts.pptx');

  out.on('error', function (err) {
    console.log(err);
  });

  pptx.generate(out);
}

async.series([
  generateCharts    // new
], finalize);