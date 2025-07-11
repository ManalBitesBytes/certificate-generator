import pandas as pd
from weasyprint import HTML, CSS
import os

def generate_certificates(csv_file, template_path, output_dir):
    """
    Generates PDF certificates from a CSV file and a single HTML template.

    Args:
        csv_file (str): Path to the CSV file containing recipient data.
        template_path (str): Path to the single HTML template file to use for all certificates.
        output_dir (str): Directory to save the generated PDF certificates.
    """
    # Create output directory if it doesn't exist
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
        with open(template_path, 'r', encoding='utf-8') as f:
            html_content_template = f.read()
        base_url_for_template = os.path.abspath(os.path.dirname(template_path))
        print(f"Successfully loaded template: {template_path} with base_url: {base_url_for_template}")

    except FileNotFoundError as e:
        print(f"Error: Template file not found: {e}")
        return
    except Exception as e:
        print(f"Error loading HTML template {template_path}: {e}")
        return

    for index, row in df.iterrows():
        recipient_name = row['Name']

        print(f"\nProcessing recipient: {recipient_name}")
        populated_html = html_content_template.replace('{{NAME}}', recipient_name)

        pdf_filename = os.path.join(output_dir, f"{recipient_name}.pdf")

        try:
            HTML(string=populated_html, base_url=base_url_for_template).write_pdf(pdf_filename)
            print(f"Successfully generated PDF: {pdf_filename}")
        except Exception as e:
            print(f"Error generating PDF for {recipient_name}: {e}")

if __name__ == "__main__":

    CSV_FILE = 'recipients/volunteers_data.csv'
    HTML_TEMPLATE = 'templates/template2.html'
    OUTPUT_FOLDER = 'pdf_volunteers'

    generate_certificates(CSV_FILE, HTML_TEMPLATE, OUTPUT_FOLDER)

    print("\nCertificate generation process completed.")
