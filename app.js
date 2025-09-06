import express from 'express';
import cron from 'node-cron';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';
import xlsx from 'xlsx';
import axios from 'axios';

// --- Configuration ---
// Load all variables from the .env file
dotenv.config();
const {
    OPENROUTER_API_KEY,
    YOUR_NAME,
    YOUR_COMPANY,
    YOUR_JOB_TITLE,
    EVENT_NAME,
    EVENT_DATE,
    POSTER_IMAGE_PATH,
    EMAIL_HOST,
    EMAIL_PORT,
    SENDER_EMAIL,
    EMAIL_PASSWORD,
    CC_EMAILS // <-- Loading the CC addresses
} = process.env;

// --- Main Email Sending Logic ---
async function sendPersonalizedEmails() {
    console.log('--- Starting Email Sending Process ---');
    
    const transporter = nodemailer.createTransport({
        host: EMAIL_HOST,
        port: parseInt(EMAIL_PORT, 10),
        secure: false, 
        auth: {
            user: SENDER_EMAIL,
            pass: EMAIL_PASSWORD,
        },
    });

    let contacts;
    try {
        const workbook = xlsx.readFile('contacts.xlsx');
        contacts = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        console.log(`Found ${contacts.length} contacts to process.`);
    } catch (error) {
        console.error("Error: Could not read contacts.xlsx. Please ensure the file exists in the same folder.");
        return;
    }

    for (const contact of contacts) {
        let { Name, Email, Designation, Persona } = contact;

        if (!Email || typeof Email !== 'string' || !Email.includes('@')) {
            console.warn(`- Skipping contact "${Name || 'Unknown'}" due to missing or invalid email address.`);
            continue;
        }

        if (!Persona && Designation) {
            console.log(`- Persona not found for ${Name}. Using Designation "${Designation}" as fallback.`);
            Persona = Designation;
        } else if (!Designation && Persona) {
            console.log(`- Designation not found for ${Name}. Using Persona "${Persona}" as fallback.`);
            Designation = Persona;
        } else if (!Persona && !Designation) {
            console.warn(`- Skipping ${Name} because both Persona and Designation are empty.`);
            continue;
        }
        
        console.log(`\nProcessing contact: ${Name} (${Email})`);
        
        const emailContent = await generateEmailContent(Name, Designation, Persona);

        if (!emailContent) {
            console.error(`- Failed to generate email content for ${Name}. Skipping.`);
            continue;
        }

        const mailOptions = {
            from: `"${YOUR_NAME}" <${SENDER_EMAIL}>`,
            to: Email,
            cc: CC_EMAILS, 
            subject: emailContent.subject,
            html: emailContent.body,
            attachments: [{
                filename: 'invitation-poster.png',
                path: POSTER_IMAGE_PATH
            }]
        };

        try {
            await transporter.sendMail(mailOptions);
            console.log(`✅ Successfully sent email to ${Name} with subject: "${emailContent.subject}"`);
        } catch (error) {
            console.error(`❌ Failed to send email to ${Name}. Error: ${error.message}`);
        }

        await new Promise(resolve => setTimeout(resolve, 5000));
    }

    console.log('\n--- Email Sending Process Finished ---');
}

// --- AI Content Generation Function (Updated with new offer logic) ---
async function generateEmailContent(name, designation, persona) {
    const prompt = `
    You are Arnav Agarwal, a Developer Advocate at Hackingly. Your tone is warm, respectful, and like a genuine peer reaching out—not a marketer.

    Your task is to generate a personalized email for ${name}, whose role is ${designation}. The goal is to invite them to our exclusive online event, "${EVENT_NAME}", on ${EVENT_DATE}.

    **Your response MUST be a valid JSON object with two keys: "subject" and "body".**

    **Instructions for the "subject" key:**
    - Create a short, compelling, and unique subject line (under 10 words).
    - It must be personalized to the recipient's role.
    - Do NOT use generic words like "Invitation," "Free," "Offer," or "Event."
    - Good examples: "A new approach to tech hiring at Decoded" for HR, or "Decoded by Hackingly | Tools for scaling your startup" for a Founder.

    **Instructions for the "body" key:**
    - The value should be a clean HTML block for the email body.
    - Follow the structure of the reference email meticulously for all personas.
    - **Greeting:** Start with "Dear ${name},". Use paragraph tags with "color:#333333;" for dark mode.
    - **Introduction & Personalization:** Combine your intro and a personalization related to their ${designation}.
    - **The Invitation:** Clearly invite them to the event.
    - **About the Event:** Briefly explain 'Decoded'.
    - **Benefits for You:** A personalized section based on their '${persona}' with a clear heading.
        - **If persona is "HR":** Describe running hiring sprints and skill assessments. Offer "1 free hiring challenge with analytics + backend support."
        - **If persona is "TPO":** Describe streamlining campus placements. Offer a "free campus hiring challenge."
        - **If persona is "Program Manager" or "College SPOC":** Describe running hackathons and bootcamps. Offer a "Free event (hackathon, bootcamp, demo day) + event planning template access."
        - **If persona is "Founder":** Describe hiring tech talent and testing ideas. Offer "One branded campaign or challenge + spotlight in Hackingly newsletter."
    - **Comprehensive Benefits Section:** A bulleted list with the heading "A Glimpse of What We Offer Across the Ecosystem:". Reorder the list to put the most relevant benefit for their persona FIRST.
        - Run hiring sprints & launch skill-based assessments.
        - Run end-to-end hackathons, bootcamps, and pitch days.
        - Hire tech talent, test product ideas, and build their brand via developer events.
        - Conduct large-scale skill assessments and streamline campus placements.
    - **Call to Action:** A styled HTML button with the link "https://www.hackingly.in/events/decoded-by-hackingly?utm=CG5BvanSGWn".
    - **Closing and Sign Off (Combined):** Combine the closing and sign-off into a single paragraph tag. Start with a unique sentence based on their persona. After the unique sentence, add two line breaks (<br><br>) followed by the full signature.
        - The full signature must be:
          Best,<br>
          ${YOUR_NAME}<br>
          ${YOUR_JOB_TITLE}<br>
          ${YOUR_COMPANY}<br>
          8290590066
        - The final HTML for a Founder must look like this: "<p style='color:#333333;'>As a fellow builder, I'm particularly excited to hear your thoughts on building innovative products. We believe the future is built together, and we'd be honored to have you be a part of it.<br><br>Best,<br>${YOUR_NAME}<br>${YOUR_JOB_TITLE}<br>${YOUR_COMPANY}<br>8290590066</p>"

    Example of a valid JSON response:
    {
      "subject": "A thought for you as a Founder",
      "body": "<p style='color:#333333;'>Dear Arnav,</p>..."
    }
    `;
    try {
        const response = await axios.post("https://openrouter.ai/api/v1/chat/completions", {
            model: "openai/gpt-3.5-turbo",
            messages: [{ role: "user", content: prompt }]
        }, { headers: { "Authorization": `Bearer ${OPENROUTER_API_KEY}` } });
        
        const content = response.data.choices[0].message.content.trim();
        // --- Parse the JSON response from the AI ---
        const parsedContent = JSON.parse(content);
        if (parsedContent.subject && parsedContent.body) {
            return parsedContent;
        }
        console.error(`- AI response for ${name} was missing subject or body.`);
        return null;

    } catch (error) {
        console.error(`- OpenRouter API request or JSON parsing failed for ${name}:`, error.message);
        return null;
    }
}


// --- Express Server and Cron Job Setup ---
const app = express();
const PORT = 3001;
let isSending = false; // Flag to prevent multiple sending jobs from running at once

app.get('/', (req, res) => {
    res.send('Email automation server is running. To send emails immediately, visit the /send-now endpoint.');
});

app.get('/send-now', (req, res) => {
    if (isSending) {
        return res.status(429).send('An email sending process is already running. Please wait for it to complete.');
    }

    res.send('Manual trigger received. Starting the email sending process now. Check your terminal for progress.');

    console.log('--- MANUAL TRIGGER ACTIVATED ---');
    isSending = true;
    sendPersonalizedEmails().finally(() => {
        isSending = false; // Reset the flag when the process is complete
        console.log('--- MANUAL PROCESS FINISHED ---');
    });
});


console.log('Scheduling the daily email job...');
cron.schedule('0 10 * * *', () => {
    if (isSending) {
        console.log('CRON JOB: Skipped because a sending process was already running.');
        return;
    }

    console.log('---------------------------------');
    console.log('CRON JOB TRIGGERED:', new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }));
    
    isSending = true;
    sendPersonalizedEmails().finally(() => {
        isSending = false; // Reset the flag when the process is complete
        console.log('--- SCHEDULED PROCESS FINISHED ---');
    });
}, {
    scheduled: true,
    timezone: "Asia/Kolkata"
});

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
    console.log(`The email sending job is scheduled to run daily at 10:00 AM IST.`);
    console.log(`To send emails immediately, open your browser and go to http://localhost:${PORT}/send-now`);
});

