import nodemailer from 'nodemailer';

class EmailHandler {
    static async sendEmail({ to, subject, body, attachments }) {
        try {
            // Configure your email service
            const transporter = nodemailer.createTransport({
                service: 'outlook', // Zet provider van de eenheidsemail hier
                auth: {
                    user: 'blank', // Vervang met Email van de eenheid
                    pass: 'blank',  // Vervang met wachtwoord van Email van de eenheid
                },
            });

            // Email options
            const mailOptions = {
                from: 'jelleswartebroekx@gmail.com', // Email zender
                to: to, // Email ontvanger
                subject: subject,
                html: body, // Email body in HTML format
                attachments: attachments, // Array voor attachments
            };

            // Stuur email
            const info = await transporter.sendMail(mailOptions);
            console.log('Email verstuurd:', info.response);
        } catch (error) {
            console.error('Fout bij versturen van mail:', error.message);
        }
    }
}

export default EmailHandler;
