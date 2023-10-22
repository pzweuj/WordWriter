## Intended Purpose
To facilitate the substitution of reserved tags within a docx template. 

This utility supports the replacement of paragraph strings, cell strings, text box strings, headers, and footers. 

Furthermore, it enables the insertion of tables and images.


### Requirements

```bash
pip install python-docx
pip install pandas
python3
```

### Basic Usage

```python
from WordWriter import WordWriter

resultsDict = {}
resultsDict["#[testheader1]#"] = "test header 1"
resultsDict["#[testheader2]#"] = "test header 2"
resultsDict["#[testString]#"] = "test strings"
resultsDict["#[testfooter]#"] = "test footer"
resultsDict["#[TX-testString2]#"] = "text box strings"
resultsDict["#[testTableString1]#"] = "cell text 1"
resultsDict["#[testTableString2]#"] = "cell text 2"
resultsDict["#[IMAGE-test1-(30,30)]#"] = "testPicture.png"
resultsDict["#[IMAGE-test2]#"] = "testPicture2.png"
resultsDict["#[IMAGE-test3-(10,10)]#"] = "testPicture.png"
resultsDict["#[TABLE-test1]#"] = "testTable.txt"

WordWriter("test.docx", "output.docx", resultsDict)
```

### Document

[Click here!](https://pzweuj.github.io/2023/10/09/WordWriter.html)

