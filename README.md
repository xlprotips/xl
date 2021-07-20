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
following will print the first 10 lines and then exit:

```bash
xlcat <path-to-xlsx> <tab-in-xlsx> -n 10
```

You could obviously do this with `head` or whatever, but this makes it slightly
easier to do without a separate tool.

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

The project is licensed under the MIT License - see the LICENSE.md_ file for
details


.. _license.md: /LICENSE.txt
