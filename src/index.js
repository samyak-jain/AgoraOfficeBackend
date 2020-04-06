const express = require('express');
const cors = require('cors');
const httpProxy = require('http-proxy');
const fileUpload = require('express-fileupload');
const asyncHandler = require('express-async-handler');
const got = require('got');
const { v4: uuidv4 } = require('uuid');

const app = express();
const corsOptions = {
    origin: true
};

const proxy = httpProxy.createProxyServer({});

proxy.on('proxyReq', asyncHandler(async(proxyReq, req, res, options) => {
    const url = 'US6-word-view.officeapps.live.com';
    proxyReq.setHeader('origin', `https://${url}`);
    proxyReq.setHeader('authority', url);
    proxyReq.setHeader('referer', `https://${url}`);
}));

proxy.on('error', asyncHandler(async(err, req, res) => {
    console.warn(err);
}));

app.use(cors(corsOptions));
app.use(fileUpload({
    createParentPath: true
}));
app.use('/files', express.static('files'));
app.use('/assets', express.static('assets'));
app.use(require('morgan')('dev', {
    skip: (req, res) => {
        return (req.url != "/" && req.url != '/do' && req.url != '/upload');
    }
}));

app.get('/', asyncHandler(async(req, res) => res.send('Hello World!')));

app.get('/do', asyncHandler(async(req, res) => {
    const mainUrl = `https://US6-word-view.officeapps.live.com/wv/wordviewerframe.aspx?embed=1&ui=en%2DUS&rs=en%2DUS&WOPISrc=http%3A%2F%2Fus6%2Dview%2Dwopi%2Ewopi%2Elive%2Enet%3A808%2Foh%2Fwopi%2Ffiles%2F%40%2FwFileId%3FwFileId%3D${encodeURIComponent(req.query.url)}&access_token_ttl=0`;
    console.info(`Sending request to URL: ${mainUrl}`);
    const response = await got(mainUrl);
    res.send(response.body);
}));

app.post('/upload', asyncHandler(async(req, res) => {
    console.log(req.files.file);

    if (!req.files || Object.keys(req.files).length === 0) {
        return res.status(400).send('No files were uploaded.');
    }

    const uploadedFile = req.files.file;
    const uploadUrl = `${uuidv4()}/${uploadedFile.name}`;

    uploadedFile.mv(`${__dirname}/files/${uploadUrl}`, (err) => {
        if (err) {
            console.log(err);
            return res.status(500).send(err);
        }

        res.send(uploadUrl);
    });
}));

app.get('/upload_ui', asyncHandler(async(req, res) => {
    res.send(`
        <!DOCTYPE HTML>
        <html lang="en">

            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <link href="./assets/dropzone.min.css" rel="stylesheet">
                <script src="./assets/dropzone.min.js"></script>
                <script>Dropzone.autoDiscover = false;</script>
                <title>File Upload</title>
            </head>
            <body>
                <div id="filePicker" class="dropzone"></div>
                <script>
                    var myDropzone = new Dropzone("div#filePicker", {
                        url: "./upload",
                        success: (file, response) => {
                            console.log(file, response);
                            const event = new CustomEvent("iframeEvent", {
                                detail: response
                            });
                            window.parent.document.dispatchEvent(event);
                        }
                    });
                </script>
            </body>
        </html>

    `)
}));

app.all('*', asyncHandler(async(req, res) => {
    const proxyUrl = 'https://US6-word-view.officeapps.live.com/wv/';

    proxy.web(req, res, {
        target: proxyUrl,
        changeOrigin: true
    });
}));

module.exports = app;