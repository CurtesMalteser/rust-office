use std::fs::File;
use std::io::Read;
use docx_rs::*;

fn heading1() -> (&'static str, Style) {
    let style_id = "Heading1";
    let style = Style::new(style_id, StyleType::Paragraph)
        .size(32)
        .bold()
        .color("0000FF")
        .text_alignment(TextAlignmentType::Center);
    (style_id, style)
}

fn body_style() -> (&'static str, Style) {
    let body_id = "Body";
    let body_style = Style::new(body_id, StyleType::Paragraph)
        .size(24);
    (body_id, body_style)
}

fn about_me() -> Table {
    // Add table.
    // Left cell with a picture of CV person
    // Right cell with a text box containing the CV information
    let mut img = File::open("../images/bot.png").unwrap();
    let mut buf = Vec::new();
    let _ = img.read_to_end(&mut buf).unwrap();
 
    // 9_525 EMUs per pixel 
    // Width: 150 * 9525 = 1_428_750 EMUs
    // Height: 100 * 9525 =   952_500 EMUs
    let pic = Pic::new(&buf).size(150 * 9525, 100 * 9525);

    let cell : Vec<TableCell> = vec![
        TableCell::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_image(pic))
            ),
        TableCell::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("I'm a bot Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur id purus et justo aliquam mollis at vitae nunc. Praesent semper condimentum euismod. Proin porttitor enim risus, sit amet sollicitudin diam eleifend at. Proin neque ligula, congue sed efficitur a, ornare tempor justo. Vestibulum eget auctor urna. Proin id massa blandit, volutpat eros a, volutpat tellus. Etiam in est et turpis consequat volutpat id blandit est. Praesent ultrices, dolor vitae congue dapibus, risus lacus fringilla diam, in malesuada diam est non arcu.").color("ffffff"))
            ).shading(Shading::new().fill("007ACC")),
    ];

    let table_rows : Vec<TableRow> = vec![TableRow::new(cell)];

    Table::new(table_rows)
        .set_grid(vec![3000, 7000])
        .indent(0)
        .align(TableAlignmentType::Center)
        .layout(TableLayoutType::Autofit)
        .clear_all_border()

}

pub fn hello() -> Result<(), DocxError> {
    let path = std::path::Path::new("./hello.docx");

    let (heading1_id, heading1) = heading1();
    let (body_id ,body_style) = body_style();
    
    let file = std::fs::File::create(path).unwrap();
    Docx::new()
        .add_table(about_me())
        .add_style(heading1)
        .add_style(body_style)
        .add_paragraph(
            Paragraph::new()
                .style(&heading1_id)
                .add_run(
                    Run::new()
                        .add_text("Hello,")
                    )
                )
        .add_paragraph(Paragraph::new()
            .style(&body_id)
            .add_run(Run::new().add_text("what's up?")))
        .build()
        .pack(file)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let result = hello();
        assert!(result.is_ok(), "Failed to create hello.docx");
    }
}
