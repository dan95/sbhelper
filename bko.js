(async () => {
    const connectionStringInput = document.getElementById("connstrtxt");
    const topicInput = document.getElementById("topictxt");
    const filterNameInput = document.getElementById("filtertxt");
    const filterValueInput = document.getElementById("filtervaltxt");
    const messageInput = document.getElementById("msgtxt");
    const sendMessageButton = document.getElementById("sendmsgbtn");
    const sendMessageForm = document.getElementById('sendmsgfrm');
    const messageTypeInput = document.getElementById('messagetypeddl');
    const sendTypeInput = document.getElementById('sendtypeddl');
    const sasTokenButton = document.getElementById("sastokenbtn");
    const sasTokenOutput = document.getElementById("sastokentxt");
    const copySasTokenButton = document.getElementById("copysasbtn");
    const sendMsgFileButton = document.getElementById('sendmsgfilebtn');
    const fileInput = document.getElementById('msgfiletxt');

    window.addEventListener('keyup', (evt) => {
        if (evt.ctrlKey && evt.code.toLowerCase() === 'enter') {
            sendMessageButton.click();
        }
    });

    const SEND_TYPE = {
        PLAIN_TEXT: 'PLAIN_TEXT',
        PLAIN_TEXT_MULTILINE: "PLAIN_TEXT_MULTILINE",
        JSON_ARRAY: 'JSON_ARRAY',
        JSON_ARRAY_PARALLEL: 'JSON_ARRAY_PARALLEL'
    };

    const __swal = window.Swal || {};
    const alerts = {
        success: (text, callback) => alerts.show('Sucesso', text, 'success', callback),
        error: (text, callback) => alerts.show('Erro', text, 'error', callback),
        warn: (text, callback) => alerts.show('Atenção', text, 'warning', callback),
        show: (title, text, icon, callback) => {
            __swal.fire({
                title: title,
                text: text,
                icon: icon,
                confirmButtonText: 'OK'
            }).then(result => {
                if (!callback || !typeof callback === 'function') {
                    return;
                }

                callback(result);
            });
        }
    };

    sasTokenButton.addEventListener('click', async () => {
        const connectionString = connectionStringInput.value;
        if (connectionString == '') {
            alerts.warn("Connection string não informada");
            return;
        }

        sasTokenOutput.value = "";
        await waitTimeout();
        const client = await sbClient(connectionString);
        sasTokenOutput.value = await client.generateSasToken();
    });

    copySasTokenButton.addEventListener('click', async () => {
        await navigator.clipboard.writeText(sasTokenOutput.value);
    });

    const applyTooltips = async (inputElements = []) => {
        const elementArray = (inputElements || []);

        await waitTimeout();

        for (let element of elementArray) {
            new bootstrap.Tooltip(element);
        }
    };

    const enableLoading = () => {
        for (let element of document.getElementsByClassName('_msgsender')) {
            element.classList.add('disabled');
        }

        for (let element of document.getElementsByClassName('_spinner')) {
            element.classList.remove('d-none');
        }
    };

    const disableLoading = () => {
        for (let element of document.getElementsByClassName('_msgsender')) {
            element.classList.remove('disabled');
        }

        for (let element of document.getElementsByClassName('_spinner')) {
            element.classList.add('d-none');
        }
    };

    const send = async (message) => {
        if (!sendMessageForm.checkValidity()) {
            sendMessageForm.reportValidity();
            return;
        }

        var client = await sbClient(connectionStringInput.value);

        if (!client.isValid) {
            alerts.warn('Connection string inválida', () => {
                connectionStringInput.focus();
            });
            return;
        }

        enableLoading();

        const messageProperties = getMessageProperties();

        switch (sendTypeInput.value) {
            case SEND_TYPE.PLAIN_TEXT:
                await sendPlainText(client, messageProperties, message); break;
            case SEND_TYPE.PLAIN_TEXT_MULTILINE:
                await sendPlainTextMultilines(client, messageProperties, message); break;
            case SEND_TYPE.JSON_ARRAY:
                await sendMessages(client, messageProperties, message);
                break;
            case SEND_TYPE.JSON_ARRAY_PARALLEL:
                await parallelSendMessages(client, messageProperties, message);
                break;
        }

        disableLoading();
    };

    sendMessageButton.addEventListener('click', async () => {
        if (messageInput.value === '') {
            alerts.warn('Informe o valor da mensagem', () => {
                connectionStringInput.focus();
            });

            return;
        }

        await send(messageInput.value);
    });

    const groupArray = (list, groupSize = 5) => {
        if (!list) {
            throw new Error('list must be an array');
        }

        if (!groupSize || groupSize <= 1) {
            throw new Error('"groupSize" must be greater than 1');
        }

        return list.reduce((a, b, i) => {
            const key = Math.floor(i / groupSize);

            if (a[key]) {
                a[key].push(b);
            } else {
                a[key] = [b];
            }

            return a;
        }, []);
    };

    const getMessageProperties = () => {
        const topic = topicInput.value;
        const contentType = messageTypeInput.value;
        let filter = {};

        if (filterNameInput.value !== "" && filterValueInput.value !== "") {
            filter[filterNameInput.value] = filterValueInput.value;
        }

        return {
            topic,
            contentType,
            filter
        };
    };

    const downloadMessageSendResult = (text) => {
        let downloadBtn = document.createElement('a');
        downloadBtn.href = URL.createObjectURL(new Blob([text], { type: 'text/plain' }));
        downloadBtn.download = `send_message_result_${crypto.randomUUID()}.txt`;

        downloadBtn.click();
    };

    const sendPlainText = async (client, messageProperties, message) => {
        const result = await client.postMessage(
            message,
            messageProperties.topic,
            messageProperties.contentType,
            messageProperties.filter);

        downloadMessageSendResult(JSON.stringify([
            result
        ], undefined, 4));
    };

    const sendMessages = async (client, messageProperties, message) => {
        const messageArray = getJsonArray(message);
        let resultList = [];

        for (let item of messageArray) {
            const itemText = JSON.stringify(item);
            const result = await client.postMessage(
                itemText,
                messageProperties.topic,
                messageProperties.contentType,
                messageProperties.filter);

            resultList.push(result);
        }

        downloadMessageSendResult(JSON.stringify(resultList, undefined, 4));
    };

    const sendPlainTextMultilines = async (client, messageProperties, message) => {
        let resultList = [];

        const messageArray = message.replaceAll('\r', '').split('\n').filter(x => x !== '');

        for (let item of messageArray) {
            const result = await client.postMessage(item,
                messageProperties.topic,
                messageProperties.contentType,
                messageProperties.filter);
            resultList.push(result);
        }

        downloadMessageSendResult(JSON.stringify(resultList, undefined, 4));
    };

    const parallelSendMessages = async (client, messageProperties, message) => {
        const groups = groupArray(
            getJsonArray(message)
        );

        let resultList = [];

        for (let group of groups) {
            var resultChunk = await Promise.all(
                group.map((item) => client.postMessage(
                    JSON.stringify(item),
                    messageProperties.topic,
                    messageProperties.contentType,
                    messageProperties.filter
                ))
            );
            resultList.push(...resultChunk);
        }

        downloadMessageSendResult(JSON.stringify(resultList, undefined, 4));
    };

    const getJsonArray = (message) => {
        try {
            const messageArray = JSON.parse(message.replaceAll('\n', '').replaceAll('\r', ''));
            if (!messageArray instanceof Array) {
                return [];
            }

            return messageArray;
        }
        catch {
            return [];
        }
    };

    sendMsgFileButton.addEventListener('click', async () => {
        if (fileInput.files.length === 0) {
            alerts.warn('Selecione um arquivo');
            return;
        }

        await send(await fileInput.files[0].text());
    });

    const waitTimeout = (timeout = 200) =>
        new Promise(r => setTimeout(r, timeout));

    const sbClient = async (connectionString) => {
        const getPropertyValue = (sourceText) =>
            sourceText.substring(sourceText.indexOf("=") + 1);

        const getHmac256Digest = async (secret, value) => {
            const encoder = new TextEncoder();

            const algorithm = {
                name: 'HMAC',
                hash: 'SHA-256'
            };

            const key = await crypto.subtle.importKey(
                'raw',
                encoder.encode(secret),
                algorithm,
                false,
                ['sign', 'verify']
            );

            const signature = await crypto.subtle.sign(
                algorithm.name,
                key,
                encoder.encode(value));

            return convertToBase64(new Uint8Array(signature));
        };

        const convertToBase64 = (bArray) => {
            let binaryString = '';
            for (let i = 0; i < bArray.length; i++) {
                binaryString += String.fromCharCode(bArray[i]);
            }

            return btoa(binaryString);
        };

        const getToken = async () => {
            let date = new Date();
            date.setHours(date.getHours() + 2);

            const epochTime = Math.floor(date.valueOf() / 1000);

            const token = await getHmac256Digest(sharedAccessKey, `${uri}\n${epochTime}`);

            return `SharedAccessSignature sr=${uri}&sig=${encodeURIComponent(token)}&se=${epochTime}&skn=${sharedAccessKeyName}`;
        };

        const postMessage = async (message, queueName, contentType, customHeaders) => {
            if (typeof message !== 'string') {
                throw Error('Message must be a string');
            }

            if (typeof contentType !== 'string' || (contentType ?? '') == '') {
                throw Error('Content type must be informed. See const FILE_CONTENT');
            }

            const token = sharedToken || await getToken();

            let headers = {
                'Content-Type': contentType,
                'Authorization': token,
                'Host': uri
            };

            if (typeof customHeaders === 'object' && !!customHeaders) {
                for (let key of Object.keys(customHeaders)) {
                    headers[key] = customHeaders[key];
                }
            }

            const response = await fetch(`https://${uri}/${queueName}/messages`, {
                headers: new Headers(headers),
                method: 'POST',
                body: message
            });

            const isSuccess = response.status === 201;

            return {
                isSuccess,
                status: response.status,
                response,
                message,
                host: uri
            };
        };

        const getMessage = async (queueName, subscription, token) => {
            const messageToken = token || await getToken();

            var response = await fetch(
                `https://${uri}/${queueName}/subscriptions/${subscription}/messages/head?timeout=100`,
                {
                    headers: new Headers({
                        'Authorization': messageToken,
                        'Host': uri,
                        'Content-Length': 0
                    }),
                    method: 'DELETE'
                }
            );

            return {
                status: response.status,
                isEmpty: response.status == 204,
                isError: !response.ok,
                message: await response.text(),
                headers: Array.from(response.headers.entries()).map(x => {
                    return {
                        key: x[0],
                        value: x[1]
                    };
                }),
                id: crypto.randomUUID()
            };
        };

        const consumeMessages = async (queueName, handleMessage, shouldContinue, subscription = 'default') => {
            if (!handleMessage || typeof handleMessage !== 'function') {
                throw new Error('Callback must be a function');
            }

            var token = await getToken();

            let brokeredMessages = [];

            while (shouldContinue === undefined || shouldContinue()) {
                console.log('[client] - fetching message...');
                const message = await getMessage(queueName, subscription, token);

                if (message.isError) {
                    continue;
                }

                if (message.isEmpty) {
                    console.log('[client] - no message available')
                    await waitTimeout();

                    continue;
                }

                await handleMessage(message);
                brokeredMessages.push(message.id);
            }
        };

        const FILE_CONTENT = {
            JSON: 'application/json;charset=utf-8',
            XML: 'text/xml;charset=utf-8'
        };

        let connectionStringParts;
        let endpoint;
        let uri;
        let sharedAccessKeyName;
        let sharedAccessKey;
        let sharedToken;

        try {
            connectionStringParts = connectionString.split(';');
            endpoint = getPropertyValue(connectionStringParts[0]);
            uri = (new URL(endpoint)).hostname;
            sharedAccessKeyName = getPropertyValue(connectionStringParts[1]);
            sharedAccessKey = getPropertyValue(connectionStringParts[2]);
            sharedToken = await getToken();

            return {
                endpoint,
                uri,
                sharedAccessKey,
                sharedAccessKeyName,
                postMessage,
                consumeMessages,
                FILE_CONTENT,
                generateSasToken: getToken,
                isValid: true
            };
        } catch (e) {
            return {
                error: e,
                isValid: false
            }
        }
    };

    await applyTooltips(
        [
            sendTypeInput,
            sasTokenButton
        ]
    );
})();