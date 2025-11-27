//! PPTX CLI - Command-line tool for creating PowerPoint presentations

use pptx_rs::cli::{Parser, Command, CreateCommand, FromMarkdownCommand, InfoCommand};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let args = &args[1..]; // Skip program name

    match Parser::parse(args) {
        Ok(Command::Create(create_args)) => {
            match CreateCommand::execute(
                &create_args.output,
                create_args.title.as_deref(),
                create_args.slides,
                create_args.template.as_deref(),
            ) {
                Ok(_) => {
                    let output = &create_args.output;
                    println!("✓ Created presentation: {output}");
                    let title = create_args.title.as_deref().unwrap_or("Presentation");
                    println!("  Title: {title}");
                    let slides = create_args.slides;
                    println!("  Slides: {slides}");
                }
                Err(e) => eprintln!("✗ Error: {e}"),
            }
        }
        Ok(Command::FromMarkdown(md_args)) => {
            match FromMarkdownCommand::execute(
                &md_args.input,
                &md_args.output,
                md_args.title.as_deref(),
            ) {
                Ok(_) => {
                    let output = &md_args.output;
                    println!("✓ Created presentation from markdown: {output}");
                    let input = &md_args.input;
                    println!("  Input: {input}");
                    let title = md_args.title.as_deref().unwrap_or("Presentation from Markdown");
                    println!("  Title: {title}");
                }
                Err(e) => eprintln!("✗ Error: {e}"),
            }
        }
        Ok(Command::Info(info_args)) => {
            match InfoCommand::execute(&info_args.file) {
                Ok(_) => {}
                Err(e) => eprintln!("✗ Error: {e}"),
            }
        }
        Ok(Command::Help) => print_help(),
        Err(e) => {
            eprintln!("✗ Error: {e}");
            print_help();
        }
    }
}

fn print_help() {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║           PPTX CLI - PowerPoint Generator                 ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();
    println!("USAGE:");
    println!("  pptx-cli <command> [options]");
    println!();
    println!("COMMANDS:");
    println!("  create <file.pptx>           Create a new presentation");
    println!("  from-md <input.md> <output>  Generate PPTX from Markdown");
    println!("  info <file.pptx>             Show presentation information");
    println!("  help                         Show this help message");
    println!();
    println!("CREATE OPTIONS:");
    println!("  --title <text>        Set presentation title");
    println!("  --slides <count>      Number of slides to create (default: 1)");
    println!("  --template <file>     Use template file");
    println!();
    println!("FROM-MD OPTIONS:");
    println!("  --title <text>        Set presentation title (default: 'Presentation from Markdown')");
    println!();
    println!("EXAMPLES:");
    println!("  # Create a simple presentation");
    println!("  pptx-cli create my.pptx");
    println!();
    println!("  # Create with custom title and slides");
    println!("  pptx-cli create my.pptx --title 'My Presentation' --slides 5");
    println!();
    println!("  # Generate from Markdown");
    println!("  pptx-cli from-md slides.md output.pptx");
    println!();
    println!("  # Generate from Markdown with custom title");
    println!("  pptx-cli from-md slides.md output.pptx --title 'My Slides'");
    println!();
    println!("  # Show file information");
    println!("  pptx-cli info my.pptx");
    println!();
    println!("MARKDOWN FORMAT:");
    println!("  # Slide Title");
    println!("  - Bullet point 1");
    println!("  - Bullet point 2");
    println!();
    println!("  # Another Slide");
    println!("  - Point A");
    println!("  - Point B");
    println!();
}
