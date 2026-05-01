//! Build a document with a header AND a footer that includes the
//! `Page X of Y` field codes.
//!
//! Run with:
//!
//! ```sh
//! cargo run --example header_footer
//! ```
//!
//! Open `header_footer.docx` in Word. The header should read
//! "Confidential — Workbench Export" on every page; the footer
//! should show the doc title on the left and "Page <n> of <total>"
//! on the right.

use docx_rust::{
    document::{num_pages_field, page_field, Footer, Header, Paragraph, Run},
    Docx, DocxResult,
};

fn main() -> DocxResult<()> {
    let mut docx = Docx::default();

    for i in 1..=40 {
        let para =
            Paragraph::default().push(Run::default().push_text(format!("Body paragraph {}.", i)));
        docx.document.push(para);
    }

    let mut header = Header::default();
    header.push(
        Paragraph::default().push(Run::default().push_text("Confidential — Workbench Export")),
    );
    docx.add_header(header);

    let mut footer = Footer::default();
    footer.push(
        Paragraph::default()
            .push(Run::default().push_text("ImageRight Workbench Export"))
            .push(Run::default().push_tab())
            .push(Run::default().push_text("Page "))
            .push(page_field())
            .push(Run::default().push_text(" of "))
            .push(num_pages_field()),
    );
    docx.add_footer(footer);

    docx.write_file("header_footer.docx")?;
    Ok(())
}
