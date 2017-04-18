# Spreadsheet Mapping Tool
Given a `.xlsx` file, gives you a graph of the cell formula dependencies. Useful for visualizing the inner workings of large spreadsheets.

![Example rendered graph](/example_output.png?raw=true "Example Graph")

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and usage purposes.

### Prerequisites
Java

Gradle (to build, otherwise grab a `.jar` from [Releases](https://github.com/haydenflinner/spreadsheet-mapping-tool/releases)

Graphvis (sudo apt install graphviz)


### Building (Skip to Using just interested in basic use)

You can import the project as an IntelliJ project, or if you prefer the command line, just

```
gradle jar
```

### Using
Run the tool from the command line and pipe the output into graphvis:

```
# java -jar excel-mapper.jar test.xlsx | dot -Tpng -o example_output.png
```


## Built With

* [Apache POI](https://poi.apache.org/) - For all of the gnarly formula parsing
* [Gradle](https://maven.apache.org/) - Dependency Management

## Contributing

Send in a pull request with whatever you'd like. Additions necessary:
* Test with non-Apache-POI-supported functions in the spreadsheet, like custom VisualBasic functions.
  * This should work fine since we're not actually evaluating any functions, only cell formula contents, but is untested.

This project is licensed under the MIT License.

