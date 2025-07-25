
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