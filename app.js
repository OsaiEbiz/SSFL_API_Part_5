import multer from 'multer';
import express from 'express';
import { createContact } from './controlls/Controller.js';

const upload = multer();
const PORT = 4000;
const app = express();

app.use(express.json());
app.use(upload.any());

try {
    app.listen(PORT, async () => {
        console.log(`Server running on PORT - ${PORT}`);
        await createContact();
    });
} catch (err) {
    console.log(`Server connection error - ${err}`);
}
