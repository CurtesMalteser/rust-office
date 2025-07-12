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

pub fn hello() -> Result<(), DocxError> {
    let path = std::path::Path::new("./hello.docx");

    let (heading1_id, heading1) = heading1();
    let (body_id ,body_style) = body_style();
    
    let file = std::fs::File::create(path).unwrap();
    Docx::new()
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

pub fn add(left: u64, right: u64) -> u64 {
    left + right
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        let _ = hello();
        assert_eq!(result, 4);
    }
}
