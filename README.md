# xl / xlcat

xlcat is like cat except for Excel files. Specifically, xlsx files (it won't
work on xls files unfortunately). It can handle *extremely large* Excel files
and will start spitting out the contents almost immediately. It is able to do
this by making some assumptions about the underlying xml and then exploiting
those assumptions via a [high-performance xml pull
parser](https://github.com/tafia/quick-xml).

xlcat takes the ideas from [sxl](https://github.com/ktr/sxl/), a Python library
that does something very similar, and puts them into a command-line app.

## Getting Started

You can download xlcat from the
[releases](https://github.com/xlprotips/xl/releases) page. Once you've
downloaded a binary for your operating system, you can use the tool to view an
Excel file with:

```bash
xlcat <path-to-xlsx> <tab-in-xlsx>
```

This will start spitting out the entire Excel file to your screen. If you have
a really big file, you may want to limit how many rows you print to screen. The
following will print the first 10 lines of the "Book1.xlsx" file included in
this repository:

```bash
$ xlcat tests/data/Book1.xlsx Sheet1 -n 10
1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18
19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36
37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54
55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72
73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90
91,92,93,94,95,2018-01-31,97,98,99,2018-02-28,101,102,103,104,105,106,107,108
109,110,111,112,113,114,115,116,117,2018-03-01,119,120,121,122,123,124,125,126
127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144
145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162
163,164,165,166,167,168,169,"Test",171,172,173,174,175,176,177,178,179,180
```

You could obviously limit the number of rows with `head` or something similar,
but this makes it slightly easier to do without a separate tool.

## xl library

If you install the Rust crate with something like:

```toml
[dependencies]
xl = "0.1.0"
```

You should be able to use the library as follows:

```rust
use xl::Workbook;

fn main () {
    let mut wb = xl::Workbook::open("tests/data/Book1.xlsx").unwrap();
    let sheets = wb.sheets();
    let sheet = sheets.get("Sheet1");
    for row in sheet.rows(&mut wb).take(5) {
        println!("{}", row);
    }
}
```

This API will likely change in the future. In particular, I do not like having
to pass the wb object in to the rows iterator, so I will probably try to find a
way to eliminate that part of the code.

You can run tests with the standard `cargo test`.

## License

The project is licensed under the MIT License - see the [License](/LICENSE.md)
file for details
