use docx_rs::{Paragraph, Run};

pub struct Experience {
    pub company: String,
    pub position: String,
    pub start_date: String,
    pub end_date: String,
    pub achievements: Vec<String>,
}

pub struct CV {
    pub about: String,
    pub image_path: String,
    pub experiences: Vec<Experience>,
}

pub trait ToDocx {
    fn to_docx(&self) -> Vec<Paragraph>;
}

impl<T: ToDocx> ToDocx for Vec<T> {
    fn to_docx(&self) -> Vec<Paragraph> {
        self.iter().flat_map(|e| e.to_docx()).collect()
    }
}

impl ToDocx for Experience {
    fn to_docx(&self) -> Vec<Paragraph> {
        let mut paragraphs = vec![
        add_position(self),
        add_date(self),
        ];
        
        // Add achievements
        paragraphs.extend(self.achievements.to_docx());
        paragraphs.push(Paragraph::new()); // Add a blank line after each experience
        
        paragraphs
    }
}

impl ToDocx for Vec<String> {
    fn to_docx(&self) -> Vec<Paragraph> {
        self.iter().map(|achievement| add_achievement(achievement)).collect()
    }
}

fn add_achievement(achievement: &str) -> Paragraph {
    Paragraph::new().add_run(Run::new().add_text(format!("• {}", achievement)))
}

fn add_position(exp: &Experience) -> Paragraph {

    let en_dash = '\u{2013}';

    Paragraph::new()
    .add_run(Run::new()
    .add_text(format!("{} {} {}", exp.position, en_dash, exp.company))
    .bold()
    .size(24))
}

fn add_date(exp: &Experience) -> Paragraph {

    let en_dash = '\u{2013}';

    Paragraph::new()
    .add_run(Run::new()
    .add_text(format!("{} {} {}", exp.start_date, en_dash, exp.end_date))
    .italic()
    .size(20)
    .color("7F7F7F"))
}