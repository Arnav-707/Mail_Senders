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
    EMAIL_PASSWORD
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
        const workbook = xlsx.readFile('Contacts.xlsx');
        contacts = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        console.log(`Found ${contacts.length} contacts to process.`);
    } catch (error) {
        console.error("Error: Could not read contacts.xlsx. Please ensure the file exists in the same folder.");
        return;
    }

    for (const contact of contacts) {
        // Simplified to only use the columns you have
        const { Name, Email, Designation, Persona } = contact;
        if (!Persona) {
            console.warn(`- Skipping ${Name} because their 'Persona' is empty.`);
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

// --- AI Content Generation Function (Updated to match the reference email) ---
async function generateEmailContent(name, designation, persona) {
    const prompt = `
    You are Arnav Agarwal, a Developer Advocate at Hackingly. Your tone is warm, respectful, and like a genuine peer reaching out—not a marketer.

    Write a personalized HTML email to ${name}. Their current role is ${designation}. The goal is to invite them to our exclusive online event, "${EVENT_NAME}", on ${EVENT_DATE}.

    **You must follow the structure of the reference email meticulously for all personas.** The output should be a clean HTML block ready to be sent.

    **Instructions for the email structure (Follow this order and style exactly):**
    1.  **Greeting:** Start with "Dear ${name}," on its own line. Example: "<p>Dear ${name},</p>"

    2.  **Introduction & Personalization:** Combine your intro and the personalization into a single paragraph. It should start with your name and title, then connect to their specific role. Example: "<p>My name is Arnav Agarwal, and I'm a Developer Advocate at Hackingly. I'm reaching out to leaders and innovators in the tech space, and your work as a ${designation} particularly stood out.</p>"

    3.  **The Invitation:** In a new paragraph, clearly invite them to the event. Example: "<p>I'd like to personally invite you to our exclusive online event, <strong>${EVENT_NAME}</strong>, happening on ${EVENT_DATE}.</p>"
    
    4.  **About the Event:** In a new paragraph, briefly explain what the event is about. Example: "<p>'Decoded' is our flagship event where we unveil the latest tools and platforms designed to help companies build, hire, and grow their tech ecosystems.</p>"

    5.  **Benefits for You (Personalized Section):** Based on their persona, '${persona}', explain the specific value *for them*. Use a clear heading like "<h4>Here’s what’s in it for you as a ${designation}:</h4>". Then, in a single paragraph, describe the pitch and offer naturally.
        - **If persona is "HR":** Describe how they can run hiring sprints and skill assessments, and mention their exclusive offer of "a free hiring challenge with full analytics and backend support."
        - **If persona is "Program Manager":** Describe how they can run hackathons and bootcamps, and mention their exclusive offer of "a free event of their choice on our platform, plus access to our event planning templates."
        - **If persona is "Founder":** Describe how they can hire tech talent and test ideas, and mention their exclusive offer of "one branded campaign or hiring challenge, plus a feature in the Hackingly newsletter."
        - **If persona is "College SPOC/TPO":** Describe how they can run skill assessments for students, and mention their exclusive offer of "one free campus hiring challenge or a skill-building hackathon for their students."

    6.  **Comprehensive Benefits Section:** After the personalized pitch, add a new section with the heading "<h4>A Glimpse of What We Offer Across the Ecosystem:</h4>". Under this heading, create a bulleted list (<ul> and <li> tags) summarizing the benefits. **Do not specify who each benefit is for.** The list should be:
        - Run hiring sprints & launch skill-based assessments.
        - Run end-to-end hackathons, bootcamps, and pitch days.
        - Hire tech talent, test product ideas, and build their brand via developer events.
        - Conduct large-scale skill assessments and streamline campus placements.

    7.  **Closing:** End with a personal closing. Example: "<p>I'd genuinely be thrilled to see you there and connect personally. We believe the future is built together, and we'd be honored to have you be a part of it.</p>"

    8.  **Sign Off:** Sign off with your name, title, and company. Example: "<p>Best,<br>${YOUR_NAME}<br>${YOUR_JOB_TITLE}<br>${YOUR_COMPANY}</p>"

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
cron.schedule('* * * * *', () => {
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

