import pandas as pd
from weasyprint import HTML, CSS
import os

def generate_certificates(csv_file, template1_path, template2_path, output_dir="generated_certificates"):
    """
    Generates PDF certificates from a CSV file and HTML templates.

    Args:
        csv_file (str): Path to the CSV file containing recipient data.
        template1_path (str): Path to the first HTML template file.
        template2_path (str): Path to the second HTML template file.
        output_dir (str): Directory to save the generated PDF certificates.
    """

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")

    try:
        # Read the CSV file
        df = pd.read_csv(csv_file)
        print(f"Successfully read CSV file: {csv_file}")
        print("CSV Data Head:")
        print(df.head())
    except FileNotFoundError:
        print(f"Error: CSV file not found at {csv_file}")
        return
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return

    try:
        with open(template1_path, 'r', encoding='utf-8') as f:
            template1_html = f.read()
        print(f"Successfully loaded template: {template1_path}")

        with open(template2_path, 'r', encoding='utf-8') as f:
            template2_html = f.read()
        print(f"Successfully loaded template: {template2_path}")

    except FileNotFoundError as e:
        print(f"Error: Template file not found: {e}")
        return
    except Exception as e:
        print(f"Error loading HTML templates: {e}")
        return


    for index, row in df.iterrows():
        recipient_name = row['name']
        recipient_template = row['template']
        recipient_email = row['email'] # Keep email for future use, not used in PDF for now

        print(f"\nProcessing recipient: {recipient_name} (Template: {recipient_template})")

        # Select the correct template
        if recipient_template == 'template1':
            html_content = template1_html
        elif recipient_template == 'template2':
            html_content = template2_html
        else:
            print(f"Warning: Unknown template '{recipient_template}' for {recipient_name}. Skipping.")
            continue

        populated_html = html_content.replace('{{NAME}}', recipient_name)
        pdf_filename = os.path.join(output_dir, f"{recipient_name}.pdf")

        try:
            HTML(string=populated_html).write_pdf(pdf_filename)
            print(f"Successfully generated PDF: {pdf_filename}")
        except Exception as e:
            print(f"Error generating PDF for {recipient_name}: {e}")

if __name__ == "__main__":

    csv_file_path = 'certified.csv'
    template1_html_path = 'templates/template1.html'
    template2_html_path = 'templates/template2.html'

    generate_certificates(csv_file_path, template1_html_path, template2_html_path)

    print("\nCertificate generation process completed.")
