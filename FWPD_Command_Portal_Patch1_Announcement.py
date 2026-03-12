from fpdf import FPDF
import os

class PDF(FPDF):
    def header(self):
        # Draw white background first
        self.set_fill_color(255, 255, 255)
        self.rect(0, 0, 210, 297, 'F')
        # Add watermark patch image (centered, faded)
        if os.path.exists('images/fwpd_patch.png'):
            # FPDF 1.x does not support opacity, so fade the image in the asset if possible
            self.image('images/fwpd_patch.png', x=30, y=40, w=150, h=150)
        # Fort Worth Police Department header
        self.set_font('Arial', 'B', 28)
        self.set_text_color(0, 0, 0)
        self.cell(0, 20, 'Fort Worth Police Department', ln=1, align='C')
        self.set_font('Arial', '', 14)
        self.cell(0, 10, 'Command Portal Team', ln=1, align='C')
        self.ln(10)

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 18)
        self.cell(0, 10, title, ln=1, align='C')
        self.ln(4)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 8, body)
        self.ln()

    def watermark(self):
        # Optionally, add a faded watermark (not supported natively by fpdf, so image is faded in the asset)
        pass

def main():
    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    # White background and watermark are now handled in header()

    pdf.chapter_title('FWPD Command Portal - Patch 1 Announcement')
    pdf.chapter_body(
        'Attention Command Staff,\n\n'
        'We are pleased to announce the release of Patch 1 for the FWPD Command Portal. Please note: due to unforeseen technical difficulties, the site\'s appearance may be temporarily affected. We ask that you disregard any visual inconsistencies as we work to resolve these issues. The core functionality and new features remain fully operational.\n\n'
        'Patch 1 - Updates & New Features:\n'
        '* An active Chat Room is now available for instant messaging among members.\n'
        '* An Event Calendar has been added to schedule and view upcoming meetings or events.\n'
        '* A direct Admin-to-Admin Message Board is now live, with direct admin messaging coming soon.\n'
        '* A Promotion Recommendation section is now available, allowing any member to recommend officers for promotion. All submissions are visible to High Command, who can review and update statuses.\n'
        '* New notification, account, and logout buttons have been added to the site header for easier access.\n'
        '* The current date and time are now displayed beneath the site header for your convenience.\n'
        '* A legal notice has been incorporated on the login page to ensure compliance and transparency.\n\n'
        'We appreciate your patience and understanding as we continue to improve the Command Portal. Please report any issues or feedback directly to the development team.\n\n'
        'Thank you for your continued support.\n\n'
        '-- FWPD Command Portal Team'
    )
    pdf.output('FWPD_Command_Portal_Patch1_Announcement.pdf')

if __name__ == '__main__':
    main()
