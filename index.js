const express = require('express');
const cors = require('cors');
// const { createProxyMiddleware } = require('http-proxy-middleware');
const fileUpload = require('express-fileupload');
const asyncHandler = require('express-async-handler');
const got = require('got');
const bodyParser = require('body-parser');
const { v4: uuidv4 } = require('uuid');

const app = express();
const corsOptions = {
    origin: true
};

// const proxy = createProxyMiddleware(
//     (path, req) => {
//         return (/^\/proxy\/(?:([^\/]+?))\/(.*)\/?$/i.test(path));
//     }, {
//     target: "https://us6-word-view.officeapps.live.com/wv",
//     changeOrigin: true,
//     router: req => {
//         return decodeURIComponent(req.path.split('/').filter(val => val)[1]);
//     },
//     pathRewrite: {
//         '^/proxy/.*/': '/'
//     },
//     onProxyReq: (proxyReq, req, res, options) => {
//         const url = options.target.hostname;  
//         // console.log(proxyReq.protocol + '://' + proxyReq.hostname + proxyReq.originalUrl);
//         proxyReq.setHeader('origin', `https://${url}`);
//         proxyReq.setHeader('authority', url);
//         proxyReq.setHeader('referer', `https://${url}`);

//         // if (req.body) {
//         //     const bodyData = JSON.stringify(req.body);
//         //     console.log("Body " + bodyData);
//         //     proxyReq.write(bodyData);
//         // }

//         // let bodyString = "";

//         // proxyReq.on("data", chunk => {
//         //     console.log("Incoming");
//         //     bodyString += chunk.toString('utf-8');
//         //     console.log(bodyString);
//         // })

//         console.log(proxyReq._host);
//         console.log(proxyReq._header);
//         console.log(proxyReq.path);
//         // console.log(proxyReq);
//     },
//     onProxyRes: proxyRes => {
//         // console.log(proxyRes);
//         console.log("Status " + proxyRes.statusCode);
//     },
//     onError: (err, req, res) => {
//         console.warn(err);
//     }
// });

const port = process.env.PORT || 3000

app.use(cors(corsOptions));
app.use(fileUpload({
    createParentPath: true
}));
app.use('/files', express.static('files'));
app.use('/assets', express.static('assets'));
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
// app.use(proxy);
app.use(require('morgan')('dev', {
    skip: (req, res) => {
        return (req.url != "/" && req.url != '/do' && req.url != '/upload');
    }
}));

app.get('/', asyncHandler(async(req, res) => res.send('Hello World!')));

app.get('/do', asyncHandler(async(req, res) => {
    const initialUrl = `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(req.query.url)}`;
    const initialResponse = await got(initialUrl);

    const re = /_iframeUrl = .+?_windowTitle/;
    const intermediateMatch = re.exec(initialResponse.body)[0];

    const interRe = /'(.+)'/;
    const finalMatch = interRe.exec(intermediateMatch);
    
    const mainUrl = decodeURIComponent(JSON.parse(`"${finalMatch[1]}"`));
    console.info(`Sending request to URL: ${mainUrl}`);
    const response = await got(mainUrl);

    const baseUrl = new URL(mainUrl);
    baseUrl.pathname = baseUrl.pathname.split("/").filter(val => val)[0];
    baseUrl.search = '';

    console.log("BaseUrl " + baseUrl.toString());

    res.send({
        baseUrl: baseUrl.toString(),
        htmlContent: response.body
    });
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

app.all(/^\/proxy\/(?:([^\/]+?))\/(.*)\/?$/i, asyncHandler(async(req, res) => {
    const url = new URL(decodeURIComponent(req.path.split('/').filter(val => val)[1]));
    const requestURL = new URL(req.path.split('/').filter(val => val).slice(2).join("/"), url);
    const searchParams = new URLSearchParams();
    for (key in req.query) {
        searchParams.append(key, req.query[key]);
    }

    const headers = req.headers;
    headers.origin = `https://${url}`;
    headers.authority = url.toString();
    headers.referer = `https://${url}`;
    // headers.host = `https://officeapps.live.com`;

    try {
        console.log(requestURL.toString());
        console.log(req.method);
        console.log(searchParams);
        console.log(req.body);
        console.log(headers);
        const response = await got(requestURL, {
            method: req.method,
            searchParams: searchParams,
            body: JSON.stringify(req.body),
            allowGetBody: true,
            headers: headers,
            rejectUnauthorized: false
        });

        res.set(response.headers);
        res.send(response.body);
    } catch (error) {
        console.error(error);
        res.writeHead(500, "Office servers are no longer working with library. Please contact developer");
    }
}));

module.exports = app;

app.listen(port, () => console.log(`Example app listening on port ${port}!`));