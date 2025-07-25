use std::fs::File;
use std::io::Read;
use std::path::Path;
use docx_rs::*;

fn about_me(image_path: impl AsRef<Path>, about: impl Into<String>) -> Table {
    // Get data for the Pic
    let mut img = File::open(image_path).unwrap();
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
            .shading(Shading::new().fill("007ACC"))
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text(about).color("ffffff"))
            ),
    ];

    let about_row = TableRow::new(cell).height_rule(HeightRule::Auto);
    let table_rows : Vec<TableRow> = vec![about_row];

    Table::new(table_rows)
        .set_grid(vec![3000, 7000])
        .indent(0)
        .align(TableAlignmentType::Center)
        .layout(TableLayoutType::Autofit)
        .clear_all_border()

}

// Got experience text from: https://enhancv.com/uk/cv-examples/software-developer/#re-resume-preview
fn experience_section() -> Vec<Paragraph> {
    vec![
        Paragraph::new()
            .add_run(Run::new().add_text("Experience").bold().size(28).color("007ACC")),

        Paragraph::new()
            .add_run(Run::new().add_text("Senior Software Developer - BankTech Innovations, Glasgow, UK").bold().size(24)),
        Paragraph::new()
            .add_run(Run::new().add_text("Feb 2017 – currently employed").italic().size(20).color("7F7F7F")),
        Paragraph::new().add_run(Run::new().add_text("• Designed and implemented an end-to-end secure payment gateway, handling over £5m in transactions weekly")),
        Paragraph::new().add_run(Run::new().add_text("• Played a key role in the development of a multi-currency exchange feature, which increased user base by 15% across European markets")),
        Paragraph::new().add_run(Run::new().add_text("• Mentored 3 junior developers, enhancing the team's technical capabilities and decreasing onboarding time by 50%")),
        Paragraph::new().add_run(Run::new().add_text("• Optimized existing codebase, which lead to a reduction in system outages by 20% and contributed to higher client retention")),
        Paragraph::new().add_run(Run::new().add_text("• Pioneered the integration of an AI-based fraud detection system, cutting down fraudulent transactions by 60%")),

        Paragraph::new(),

        Paragraph::new()
            .add_run(Run::new().add_text("Software Developer - Global Finance Tech, Edinburgh, UK").bold().size(24)),
        Paragraph::new()
            .add_run(Run::new().add_text("Sep 2014 – Feb 2017").italic().size(20).color("7F7F7F")),
        Paragraph::new().add_run(Run::new().add_text("• Implemented new features for an investment tracking platform, resulting in a user growth of 10% within the first quarter")),
        Paragraph::new().add_run(Run::new().add_text("• Collaborated on a cross-functional team to redesign the user interface, improving user experience scores by 25%")),
        Paragraph::new().add_run(Run::new().add_text("• Contributed to the reduction of load times for high-traffic applications by optimizing code, achieving a performance boost of 15%")),
        Paragraph::new().add_run(Run::new().add_text("• Resolved critical bugs, ensuring 99% system uptime and enhancing customer trust and business continuity")),
    ]
}


pub fn hello(about: impl Into<String>) -> Result<(), DocxError> {
    let path = std::path::Path::new("./hello.docx");
    
    let file = std::fs::File::create(path).unwrap();
    
    let image_path = "../images/bot.png".to_string();

    let paragraphs = experience_section();

    let cv_docx = Docx::new()
        .add_table(about_me(image_path, about))
        .add_paragraph(Paragraph::new().line_spacing(LineSpacing::new().before(200).after(200)))
        .add_paragraph(Paragraph::new());


    let cv_docx = paragraphs.into_iter().fold(cv_docx, |doc, para| {
        doc.add_paragraph(para)
    });

    cv_docx.build().pack(file)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let  about = "I'm a bot Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur id purus et justo aliquam mollis at vitae nunc. Praesent semper condimentum euismod. Proin porttitor enim risus, sit amet sollicitudin diam eleifend at. Proin neque ligula, congue sed efficitur a, ornare tempor justo. Vestibulum eget auctor urna. Proin id massa blandit, volutpat eros a, volutpat tellus. Etiam in est et turpis consequat volutpat id blandit est. Praesent ultrices, dolor vitae congue dapibus, risus lacus fringilla diam, in malesuada diam est non arcu.".to_string();
        let result = hello(about);
        assert!(result.is_ok(), "Failed to create hello.docx");
    }
}
