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

        // --- Change 1: Skip if Email is missing ---
        if (!Email || typeof Email !== 'string' || !Email.includes('@')) {
            console.warn(`- Skipping contact "${Name || 'Unknown'}" due to missing or invalid email address.`);
            continue;
        }

        // --- Change 2: Fallback logic for Persona and Designation ---
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
        
        const emailHtmlBody = await generateEmailContent(Name, Designation, Persona);

        if (!emailHtmlBody) {
            console.error(`- Failed to generate email content for ${Name}. Skipping.`);
            continue;
        }

        const mailOptions = {
            from: `"${YOUR_NAME}" <${SENDER_EMAIL}>`,
            to: Email,
            // --- Change 3: Add the CC email addresses ---
            cc: CC_EMAILS, 
            subject: `An Invitation to ${EVENT_NAME}`,
            html: emailHtmlBody,
            attachments: [{
                filename: 'invitation-poster.png',
                path: POSTER_IMAGE_PATH
            }]
        };

        try {
            await transporter.sendMail(mailOptions);
            console.log(`✅ Successfully sent email to ${Name}.`);
        } catch (error) {
            console.error(`❌ Failed to send email to ${Name}. Error: ${error.message}`);
        }

        await new Promise(resolve => setTimeout(resolve, 5000));
    }

    console.log('\n--- Email Sending Process Finished ---');
}

// --- AI Content Generation Function (Updated with fixes) ---
async function generateEmailContent(name, designation, persona) {
    const prompt = `
    You are Arnav Agarwal, a Developer Advocate at Hackingly. Your tone is warm, respectful, and like a genuine peer reaching out—not a marketer.

    Write a personalized HTML email to ${name}. Their current role is ${designation}. The goal is to invite them to our exclusive online event, "${EVENT_NAME}", on ${EVENT_DATE}.

    **You must follow the structure of the reference email meticulously for all personas.** The output should be a clean HTML block ready to be sent.

    **Instructions for the email structure (Follow this order and style exactly):**
    1.  **Greeting:** Start with "Dear ${name}," on its own line. Use a paragraph tag with a dark text color for dark mode compatibility. Example: "<p style="color:#333333;">Dear ${name},</p>"

    2.  **Introduction & Personalization:** Combine your intro and personalization into a single paragraph. Apply the dark text color. Example: "<p style="color:#333333;">My name is Arnav Agarwal, and I'm a Developer Advocate at Hackingly. I'm reaching out to leaders and innovators in the tech space, and your work as a ${designation} particularly stood out.</p>"

    3.  **The Invitation:** In a new paragraph, clearly invite them to the event. Apply the dark text color. Example: "<p style="color:#333333;">I'd like to personally invite you to our exclusive online event, <strong>${EVENT_NAME}</strong>, happening on ${EVENT_DATE}.</p>"
    
    4.  **About the Event:** In a new paragraph, briefly explain what the event is about. Apply the dark text color. Example: "<p style="color:#333333;">'Decoded' is our flagship event where we unveil the latest tools and platforms designed to help companies build, hire, and grow their tech ecosystems.</p>"

    5.  **Benefits for You (Personalized Section):** Based on their persona, '${persona}', explain the specific value *for them*. Use a clear heading like "<h4 style="color:#111111;">Here’s what’s in it for you as a ${designation}:</h4>". Then, in a single paragraph with the dark text color, describe the pitch and offer naturally.
        - **If persona is "HR":** Describe how they can run hiring sprints and skill assessments, and mention their exclusive offer of "a free hiring challenge with full analytics and backend support."
        - **If persona is "Program Manager":** Describe how they can run hackathons and bootcamps, and mention their exclusive offer of "a free event of their choice on our platform, plus access to our event planning templates."
        - **If persona is "Founder":** Describe how they can hire tech talent and test ideas, and mention their exclusive offer of "one branded campaign or hiring challenge, plus a feature in the Hackingly newsletter."
        - **If persona is "College SPOC/TPO":** Describe how they can run skill assessments for students, and mention their exclusive offer of "one free campus hiring challenge or a skill-building hackathon for their students."

    6.  **Comprehensive Benefits Section:** After the personalized pitch, add a new section with the heading "<h4 style="color:#111111;">A Glimpse of What We Offer Across the Ecosystem:</h4>". Under this heading, create a bulleted list (<ul> and <li style="color:#333333;"> tags). **Crucially, make the list item that is most relevant to the recipient's persona bold by wrapping it in <strong> tags.**
        - For an HR persona, the list would be: <li><strong>Run hiring sprints & launch skill-based assessments.</strong></li><li>Run end-to-end hackathons, bootcamps, and pitch days.</li><li>Hire tech talent, test product ideas, and build their brand via developer events.</li><li>Conduct large-scale skill assessments and streamline campus placements.</li>
        - For a Founder persona, the third item would be bold, etc.

    7.  **Call to Action:** After the benefits list, add a clear call to action with the registration link. It should be a styled HTML button. The link is "https://www.hackingly.in/events/decoded-by-hackingly?utm=CG5BvanSGWn".
        The HTML for the button must be:
        "<p style='text-align:center; margin: 30px 0;'><a href='https://www.hackingly.in/events/decoded-by-hackingly?utm=CG5BvanSGWn' target='_blank' style='background-color: #007bff; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;'>Register for Decoded</a></p>"

    8.  **Closing:** End with a personal closing. Apply the dark text color. Example: "<p style="color:#333333;">I'd genuinely be thrilled to see you there and connect personally. We believe the future is built together, and we'd be honored to have you be a part of it.</p>"

    9.  **Sign Off:** Sign off with your name, title, and company. Apply the dark text color. Example: "<p style="color:#333333;">Best,<br>${YOUR_NAME}<br>${YOUR_JOB_TITLE}<br>${YOUR_COMPANY}</p>"

    Format the entire response as a single block of HTML. Do not include a P.S., <html>, <body>, or any image tags. The poster will be a separate attachment.
    `;
    try {
        const response = await axios.post("https://openrouter.ai/api/v1/chat/completions", {
            model: "openai/gpt-3.5-turbo",
            messages: [{ role: "user", content: prompt }]
        }, { headers: { "Authorization": `Bearer ${OPENROUTER_API_KEY}` } });
        return response.data.choices[0].message.content.trim();
    } catch (error) {
        console.error(`- OpenRouter API request failed for ${name}:`, error.response ? error.response.data.error.message : "An unknown API error occurred");
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

