Office.onReady(() => {
    document.getElementById("generate-button").onclick = generatePPT;

    document.getElementById("ai-service-select").addEventListener('change', function() {
        const azureOpenaiConfig = document.getElementById("azure-openai-config");
        const openaiApiConfig = document.getElementById("openai-api-config");
        const customApiConfig = document.getElementById("custom-api-config");
        if (this.value === 'azure-openai') {
            azureOpenaiConfig.style.display = 'block';
            openaiApiConfig.style.display = 'none';
            customApiConfig.style.display = 'none';
        } else if (this.value === 'openai-api') {
            azureOpenaiConfig.style.display = 'none';
            openaiApiConfig.style.display = 'block';
            customApiConfig.style.display = 'none';
        } else if (this.value === 'custom') {
            azureOpenaiConfig.style.display = 'none';
            openaiApiConfig.style.display = 'none';
            customApiConfig.style.display = 'block';
        } else {
            azureOpenaiConfig.style.display = 'none';
            openaiApiConfig.style.display = 'none';
            customApiConfig.style.display = 'none';
        }
    });
});

async function generatePPT() {
    const prompt = document.getElementById("prompt-input").value;
    if (prompt) {
        // 调用AI服务生成PPT内容
        const aiOutput = await callAIService(prompt);
        document.getElementById("output-text").innerText = aiOutput;
        // 将AI生成的内容插入到PPT中
        insertSlide(aiOutput);
    } else {
        document.getElementById("output-text").innerText = "请输入PPT页面要求。";
    }
}

async function callAIService(prompt) {
    const aiService = document.getElementById("ai-service-select").value;
    let aiOutput = "";

    if (aiService === 'azure-openai') {
        // TODO: 调用 Azure OpenAI Service API
        aiOutput = await callAzureOpenAIService(prompt);
    } else if (aiService === 'openai-api') {
        // TODO: 调用 OpenAI API
        aiOutput = await callOpenAIService(prompt);
    } else if (aiService === 'custom') {
        // TODO: 调用 Custom API
        const apiUrl = document.getElementById("custom-api-url").value;
        const apiKey = document.getElementById("custom-api-key").value;
        aiOutput = await callCustomAIService(prompt, apiUrl, apiKey);
    }

    return aiOutput;
}

async function callAzureOpenAIService(prompt) {
    const apiKey = document.getElementById("azure-openai-api-key").value;
    const apiEndpoint = document.getElementById("azure-openai-endpoint").value;
    const deploymentName = document.getElementById("azure-openai-deployment-name").value;

    if (!apiKey || !apiEndpoint || !deploymentName) {
        return "请先输入 Azure OpenAI API Key, Endpoint 和 Deployment Name。";
    }

    const apiUrl = `${apiEndpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2023-05-15`; // Azure OpenAI API endpoint

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'api-key': apiKey // Azure OpenAI 使用 'api-key' header
            },
            body: JSON.stringify({
                messages: [{ role: "user", content: prompt }],
                max_tokens: 200,
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const aiOutput = data.choices[0].message.content;
        return aiOutput;

    } catch (error) {
        console.error("Error calling Azure OpenAI API:", error);
        return `调用Azure OpenAI API失败: ${error.message}`;
    }
}

async function callOpenAIService(prompt) {
    const apiKey = document.getElementById("openai-api-key").value;
    if (!apiKey) {
        return "请先输入OpenAI API Key。";
    }

    const apiUrl = "https://api.openai.com/v1/chat/completions"; // 使用chat completions API
    const model = "gpt-3.5-turbo"; // 选择合适的模型

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: model,
                messages: [{ role: "user", content: prompt }], // 使用messages参数
                max_tokens: 200, // 可以根据需要调整
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        // chat completions API 返回 choices 数组
        const aiOutput = data.choices[0].message.content; 
        return aiOutput;

    } catch (error) {
        console.error("Error calling OpenAI API:", error);
        return `调用OpenAI API失败: ${error.message}`;
    }
}

async function callCustomAIService(prompt, apiUrl, apiKey) {
    const customApiUrl = document.getElementById("custom-api-url").value;
    const customApiKey = document.getElementById("custom-api-key").value;

    if (!customApiUrl) {
        return "请先输入 Custom API URL。";
    }

    try {
        const response = await fetch(customApiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${customApiKey}` // 假设 Custom API 使用 Bearer Token 认证，可以根据实际情况修改
            },
            body: JSON.stringify({
                prompt: prompt // 假设 Custom API 接收 prompt 参数，可以根据实际情况修改
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        // 假设 Custom API 返回的 JSON 数据中包含 text 字段，可以根据实际情况修改
        const aiOutput = data.text; 
        return aiOutput;

    } catch (error) {
        console.error("Error calling Custom API:", error);
        return `调用Custom API失败: ${error.message}`;
    }
}

function insertSlide(content) {
    Office.context.document.setSelectedDataAsync(content, {
        coercionType: Office.CoercionType.Text
    }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('Error: ' + asyncResult.error.message);
        }
    });
}
