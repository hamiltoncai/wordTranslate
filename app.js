document.addEventListener('DOMContentLoaded', () => {
    // 获取DOM元素
    const modelTypeSelect = document.getElementById('model-type');
    const openaiConfig = document.getElementById('openai-config');
    const ollamaConfig = document.getElementById('ollama-config');
    const testConnectionBtn = document.getElementById('test-connection');
    const connectionStatus = document.getElementById('connection-status');
    const uploadBtn = document.getElementById('upload-btn');
    const fileInput = document.getElementById('file-input');
    const translateBtn = document.getElementById('translate-btn');
    const saveBtn = document.getElementById('save-btn');
    const originalContent = document.getElementById('original-content');
    const translatedContent = document.getElementById('translated-content');
    
    // 存储原始文档内容和结构
    let originalDocumentContent = [];
    let originalDocumentStructure = null;
    let currentFile = null;
    
    // 切换模型类型显示对应的配置
    modelTypeSelect.addEventListener('change', () => {
        if (modelTypeSelect.value === 'openai') {
            openaiConfig.style.display = 'block';
            ollamaConfig.style.display = 'none';
        } else {
            openaiConfig.style.display = 'none';
            ollamaConfig.style.display = 'block';
        }
    });
    
    // 测试大模型连接
    testConnectionBtn.addEventListener('click', async () => {
        connectionStatus.textContent = '正在测试连接...';
        connectionStatus.className = '';
        
        try {
            const isConnected = await testModelConnection();
            if (isConnected) {
                connectionStatus.textContent = '连接成功！';
                connectionStatus.className = 'success';
            } else {
                connectionStatus.textContent = '连接失败，请检查配置。';
                connectionStatus.className = 'error';
            }
        } catch (error) {
            console.error('连接测试错误:', error);
            connectionStatus.textContent = `连接错误: ${error.message}`;
            connectionStatus.className = 'error';
        }
    });
    
    // 测试模型连接
    async function testModelConnection() {
        const modelType = modelTypeSelect.value;
        
        if (modelType === 'openai') {
            const apiKey = document.getElementById('openai-api-key').value;
            const model = document.getElementById('openai-model').value;
            
            if (!apiKey) {
                throw new Error('请输入OpenAI API Key');
                apiKey= "2c02a5f2-59e3-4759-92c0-7bc997e10183";
            }
            
            try {
                const response = await fetch('https://api.openai.com/v1/chat/completions', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${apiKey}`
                    },
                    body: JSON.stringify({
                        model: model,
                        messages: [{ role: 'user', content: 'Hello' }],
                        max_tokens: 5
                    })
                });
                
                const data = await response.json();
                return response.ok && data.choices && data.choices.length > 0;
            } catch (error) {
                console.error('OpenAI API 错误:', error);
                return false;
            }
        } else if (modelType === 'ollama') {
            const endpoint = document.getElementById('ollama-endpoint').value;
            const model = document.getElementById('ollama-model').value;
            
            if (!endpoint || !model) {
                throw new Error('请输入Ollama端点和模型名称');
            }
            
            try {
                // 确保使用正确的API端点路径
                const apiUrl = `${endpoint}/api/generate`;
                const response = await fetch(apiUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        model: model,
                        prompt: 'Hello',
                        stream: false
                    })
                });
                
                const data = await response.json();
                // Ollama API可能返回不同格式的响应，检查多种可能的成功情况
                return response.ok && (data.response !== undefined || data.model !== undefined || data.created_at !== undefined);
            } catch (error) {
                console.error('Ollama API 错误:', error);
                return false;
            }
        }
        
        return false;
    }
    
    // 上传文件按钮点击事件
    uploadBtn.addEventListener('click', () => {
        fileInput.click();
    });
    
    // 文件选择事件
    fileInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        
        if (file.name.endsWith('.docx')) {
            currentFile = file;
            try {
                // 清空内容区域
                originalContent.innerHTML = '<div class="loading">正在加载文档...</div>';
                translatedContent.innerHTML = '';
                
                // 使用mammoth.js解析Word文档
                const arrayBuffer = await file.arrayBuffer();
                // 配置mammoth.js以保留更多格式信息
                const result = await mammoth.convertToHtml({
                    arrayBuffer,
                    styleMap: [
                        "p[style-name='Heading 1'] => h1:fresh",
                        "p[style-name='Heading 2'] => h2:fresh",
                        "p[style-name='Heading 3'] => h3:fresh",
                        "p[style-name='Heading 4'] => h4:fresh",
                        "p[style-name='Heading 5'] => h5:fresh",
                        "p[style-name='Heading 6'] => h6:fresh",
                        // 直接映射常见颜色，确保正确应用
                        "r[color='FF0000'] => span.red:fresh",
                        "r[color='0000FF'] => span.blue:fresh",
                        "r[color='008000'] => span.color-green:fresh",
                        "r[color='800080'] => span.color-purple:fresh",
                        "r[color='FFFF00'] => span.color-yellow:fresh",
                        "r[color='FFA500'] => span.color-orange:fresh",
                        "r[color='000000'] => span.color-black:fresh",
                        // 修改颜色映射，直接使用color-前缀类
                        "r[color='red'] => span.red:fresh",
                        "r[color='blue'] => span.blue:fresh",
                        "r[color] => span.color-$value:fresh",
                        // 确保字体样式正确应用
                        "r[bold] => span.bold:fresh",
                        "r[italic] => span.italic:fresh",
                        "r[underline] => span.underline:fresh",
                        // 确保字号正确应用
                        "r[font-size='8'] => span.size-8:fresh",
                        "r[font-size='9'] => span.size-9:fresh",
                        "r[font-size='10'] => span.size-10:fresh",
                        "r[font-size='11'] => span.size-11:fresh",
                        "r[font-size='12'] => span.size-12:fresh",
                        "r[font-size='14'] => span.size-14:fresh",
                        "r[font-size='16'] => span.size-16:fresh",
                        "r[font-size='18'] => span.size-18:fresh",
                        "r[font-size='20'] => span.size-20:fresh",
                        "r[font-size='22'] => span.size-22:fresh",
                        "r[font-size='24'] => span.size-24:fresh",
                        "r[font-size='26'] => span.size-26:fresh",
                        "r[font-size='28'] => span.size-28:fresh",
                        "r[font-size='36'] => span.size-36:fresh",
                        "r[font-size='48'] => span.size-48:fresh",
                        "r[font-size='72'] => span.size-72:fresh",
                        "r[font-size] => span.size-$value:fresh",
                        "r[font] => span.font-$value:fresh"
                    ],
                    transformDocument: mammoth.transforms.paragraph(function(paragraph) {
                        // 提取段落样式信息
                        const properties = paragraph.properties || {};
                        const style = properties.style || {};
                        const styleId = style.styleId || "";
                        
                        // 处理段落中的文本运行，提取更多格式信息
                        if (paragraph.children) {
                            paragraph.children = paragraph.children.map(function(child) {
                                if (child.type === 'run' && child.properties) {
                                    // 提取并保存更多的格式属性
                                    if (child.properties.color) {
                                        // 处理常见颜色，确保正确映射
                                        const color = child.properties.color.toUpperCase();
                                        if (color === 'FF0000') {
                                            child.properties.colorClass = 'red';
                                        } else if (color === '0000FF') {
                                            child.properties.colorClass = 'blue';
                                        } else if (color === '008000') {
                                            child.properties.colorClass = 'color-green';
                                        } else if (color === '800080') {
                                            child.properties.colorClass = 'color-purple';
                                        } else if (color === 'FFFF00') {
                                            child.properties.colorClass = 'color-yellow';
                                        } else if (color === 'FFA500') {
                                            child.properties.colorClass = 'color-orange';
                                        } else if (color === '000000') {
                                            child.properties.colorClass = 'color-black';
                                        } else {
                                            // 其他颜色使用通用格式
                                            child.properties.colorClass = 'color-' + color;
                                        }
                                        
                                        // 记录颜色信息，便于调试
                                        console.log('提取到颜色:', color, '映射为类:', child.properties.colorClass);
                                    }
                                    
                                    if (child.properties.fontSize) {
                                        // 确保字号信息被正确提取和映射
                                        const fontSize = child.properties.fontSize;
                                        child.properties.fontSizeClass = 'size-' + fontSize;
                                        
                                        // 记录字号信息，便于调试
                                        console.log('提取到字号:', fontSize, '映射为类:', child.properties.fontSizeClass);
                                    }
                                    
                                    if (child.properties.font) {
                                        child.properties.fontClass = 'font-' + child.properties.font.replace(/\s+/g, '-');
                                    }
                                }
                                return child;
                            });
                        }
                        return paragraph;
                    })
                });
                
                // 保存原始文档的ArrayBuffer以便后续处理
                originalDocumentStructure = arrayBuffer;
                
                // 显示解析后的HTML内容
                originalContent.innerHTML = result.value;
                
                // 将文档内容分段存储，以便翻译
                parseDocumentContent();
                
                // 启用翻译按钮
                translateBtn.disabled = false;
                saveBtn.disabled = true;
            } catch (error) {
                console.error('文档解析错误:', error);
                originalContent.innerHTML = `<div class="error">文档解析错误: ${error.message}</div>`;
            }
        } else {
            alert('请上传.docx格式的Word文档');
        }
    });
    
    // 解析文档内容为可翻译的段落
    function parseDocumentContent() {
        originalDocumentContent = [];
        
        // 获取所有段落、标题、列表项等元素
        const elements = originalContent.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td, th');
        
        elements.forEach((element, index) => {
            // 为每个元素添加数据属性，以便后续匹配翻译结果
            element.setAttribute('data-index', index);
            
            // 提取元素的样式信息
            const styles = window.getComputedStyle(element);
            const styleInfo = {
                color: styles.color,
                fontFamily: styles.fontFamily,
                fontSize: styles.fontSize,
                fontWeight: styles.fontWeight,
                fontStyle: styles.fontStyle,
                textDecoration: styles.textDecoration,
                backgroundColor: styles.backgroundColor
            };
            
            // 检查元素内是否有带样式的子元素
            const styledChildren = [];
            element.querySelectorAll('span').forEach(span => {
                // 获取span的类名，这些类名可能包含mammoth.js转换的样式信息
                const classList = Array.from(span.classList);
                const spanStyles = window.getComputedStyle(span);
                
                // 详细记录每个span的类名和样式，便于调试
                console.log('Span类名:', classList, 'Span内容:', span.textContent);
                
                styledChildren.push({
                    text: span.textContent,
                    classList: classList,
                    className: span.className, // 保存完整的className字符串
                    styles: {
                        color: spanStyles.color,
                        fontFamily: spanStyles.fontFamily,
                        fontSize: spanStyles.fontSize,
                        fontWeight: spanStyles.fontWeight,
                        fontStyle: spanStyles.fontStyle,
                        textDecoration: spanStyles.textDecoration
                    },
                    outerHTML: span.outerHTML // 保存完整的HTML，包括标签和属性
                });
            });
            
            // 记录整个元素的HTML结构，便于调试
            console.log('元素HTML:', element.outerHTML);
            
            // 存储元素内容、类型和样式信息
            originalDocumentContent.push({
                index,
                type: element.tagName.toLowerCase(),
                content: element.textContent.trim(),
                element,
                styleInfo,
                styledChildren,
                html: element.innerHTML // 保存HTML内容以保留格式
            });
        });
    }
    
    // 翻译按钮点击事件
    translateBtn.addEventListener('click', async () => {
        if (originalDocumentContent.length === 0) {
            alert('请先上传Word文档');
            return;
        }
        
        try {
            // 清空翻译区域并显示加载提示
            translatedContent.innerHTML = '<div class="loading">正在翻译文档...</div>';
            
            // 获取目标语言
            const targetLanguage = document.getElementById('target-language').value;
            
            // 批量翻译文档内容
            await translateDocument(targetLanguage);
            
            // 启用保存按钮
            saveBtn.disabled = false;
        } catch (error) {
            console.error('翻译错误:', error);
            translatedContent.innerHTML = `<div class="error">翻译错误: ${error.message}</div>`;
        }
    });
    
    // 翻译文档内容
    async function translateDocument(targetLanguage) {
        // 准备翻译内容
        const batchSize = 3; // 减小每批翻译的段落数，提高成功率
        const batches = [];
        
        // 将内容分批，以避免请求过大
        for (let i = 0; i < originalDocumentContent.length; i += batchSize) {
            const batch = originalDocumentContent.slice(i, i + batchSize);
            batches.push(batch);
        }
        
        // 创建翻译结果容器
        translatedContent.innerHTML = '';
        const translatedElements = {};
        
        // 逐批翻译
        for (let i = 0; i < batches.length; i++) {
            const batch = batches[i];
            const batchTexts = batch.map(item => item.content).filter(text => text.trim() !== '');
            
            if (batchTexts.length === 0) continue;
            
            // 更新进度
            const progressPercent = Math.round((i / batches.length) * 100);
            translatedContent.innerHTML = `<div class="loading">正在翻译文档... ${progressPercent}%</div>`;
            
            try {
                console.log(`开始翻译批次 ${i+1}/${batches.length}，包含 ${batchTexts.length} 个段落`);
                console.log('待翻译文本:', batchTexts);
                
                // 调用翻译API
                const translatedTexts = await translateTexts(batchTexts, targetLanguage);
                
                console.log('翻译结果:', translatedTexts);
                console.log(`翻译结果数量: ${translatedTexts.length}, 原文数量: ${batchTexts.length}`);
                
                // 确保翻译结果与原文数量匹配
                if (translatedTexts.length < batchTexts.length) {
                    console.warn('翻译结果数量少于原文数量，将使用原文填充缺失部分');
                    // 填充缺失的翻译结果
                    while (translatedTexts.length < batchTexts.length) {
                        translatedTexts.push(`[翻译失败: ${batchTexts[translatedTexts.length]}]`);
                    }
                }
                
                // 将翻译结果与原始元素匹配
                let textIndex = 0;
                batch.forEach(item => {
                    if (item.content.trim() !== '') {
                        // 创建对应的翻译元素
                        const translatedElement = document.createElement(item.type);
                        
                        // 确保有对应的翻译结果
                        const translatedText = textIndex < translatedTexts.length ? 
                            translatedTexts[textIndex] : 
                            `[翻译失败: ${item.content}]`;
                        
                        // 应用原始元素的样式
                        if (item.styleInfo) {
                            translatedElement.style.color = item.styleInfo.color;
                            translatedElement.style.fontFamily = item.styleInfo.fontFamily;
                            translatedElement.style.fontSize = item.styleInfo.fontSize;
                            translatedElement.style.fontWeight = item.styleInfo.fontWeight;
                            translatedElement.style.fontStyle = item.styleInfo.fontStyle;
                            translatedElement.style.textDecoration = item.styleInfo.textDecoration;
                            translatedElement.style.backgroundColor = item.styleInfo.backgroundColor;
                        }
                        
                        // 复制原始元素的类名，确保样式正确应用
                        if (item.element && item.element.classList) {
                            Array.from(item.element.classList).forEach(className => {
                                translatedElement.classList.add(className);
                            });
                        }
                        
                        // 如果原始元素有带样式的子元素，需要保留这些样式
                        if (item.styledChildren && item.styledChildren.length > 0) {
                            // 使用原始HTML结构作为模板
                            let htmlContent = item.html;
                            
                            // 记录原始样式信息，用于调试
                            console.log('原始样式子元素:', item.styledChildren);
                            
                            // 如果有多个带样式的子元素，尝试智能分配翻译文本
                            if (item.styledChildren.length > 1) {
                                // 创建一个临时容器来保存翻译后的HTML
                                const tempContainer = document.createElement('div');
                                tempContainer.innerHTML = item.html;
                                
                                // 获取所有文本节点和带样式的span元素
                                const textNodes = [];
                                const walkNodes = (node) => {
                                    if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
                                        textNodes.push(node);
                                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                                        if (node.tagName.toLowerCase() === 'span') {
                                            textNodes.push(node);
                                        } else {
                                            Array.from(node.childNodes).forEach(walkNodes);
                                        }
                                    }
                                };
                                Array.from(tempContainer.childNodes).forEach(walkNodes);
                                
                                // 简单地将翻译文本分配给第一个文本节点
                                // 注意：这是一个简化处理，实际应用中可能需要更复杂的文本分配算法
                                if (textNodes.length > 0) {
                                    if (textNodes[0].nodeType === Node.TEXT_NODE) {
                                        textNodes[0].textContent = translatedText;
                                    } else {
                                        // 保留span元素的类名和样式属性，只替换文本内容
                                        const originalSpan = textNodes[0];
                                        originalSpan.textContent = translatedText;
                                    }
                                    translatedElement.innerHTML = tempContainer.innerHTML;
                                } else {
                                    // 如果没有找到文本节点，直接使用翻译文本
                                    translatedElement.innerHTML = translatedText;
                                }
                            } else if (item.styledChildren.length === 1) {
                                // 只有一个带样式的子元素，保留其样式信息
                                const styledChild = item.styledChildren[0];
                                
                                // 创建一个新的span元素，保留原始样式
                                const tempSpan = document.createElement('span');
                                
                                // 复制原始span的所有类名
                                if (styledChild.classList && styledChild.classList.length > 0) {
                                    styledChild.classList.forEach(className => {
                                        tempSpan.classList.add(className);
                                    });
                                } else if (styledChild.className) {
                                    tempSpan.className = styledChild.className;
                                }
                                
                                // 应用内联样式
                                if (styledChild.styles) {
                                    if (styledChild.styles.color) tempSpan.style.color = styledChild.styles.color;
                                    if (styledChild.styles.fontFamily) tempSpan.style.fontFamily = styledChild.styles.fontFamily;
                                    if (styledChild.styles.fontSize) tempSpan.style.fontSize = styledChild.styles.fontSize;
                                    if (styledChild.styles.fontWeight) tempSpan.style.fontWeight = styledChild.styles.fontWeight;
                                    if (styledChild.styles.fontStyle) tempSpan.style.fontStyle = styledChild.styles.fontStyle;
                                    if (styledChild.styles.textDecoration) tempSpan.style.textDecoration = styledChild.styles.textDecoration;
                                }
                                
                                // 设置翻译后的文本
                                tempSpan.textContent = translatedText;
                                
                                // 将新的span元素添加到翻译元素中
                                translatedElement.innerHTML = tempSpan.outerHTML;
                                
                                // 记录应用的样式信息，便于调试
                                console.log('应用样式到翻译元素:', tempSpan.outerHTML);
                            } else {
                                // 没有带样式的子元素，直接设置文本
                                translatedElement.innerHTML = translatedText;
                            }
                        } else {
                            // 没有特殊格式，直接设置文本
                            translatedElement.textContent = translatedText;
                        }
                        
                        translatedElement.setAttribute('data-original-index', item.index);
                        
                        // 存储翻译元素
                        translatedElements[item.index] = translatedElement;
                        textIndex++;
                    }
                });
            } catch (error) {
                console.error(`批次 ${i+1} 翻译错误:`, error);
                
                // 不立即抛出错误，而是为这个批次的所有段落创建错误提示
                batch.forEach(item => {
                    if (item.content.trim() !== '') {
                        const errorElement = document.createElement(item.type);
                        errorElement.textContent = `[翻译错误: ${error.message}]`;
                        errorElement.setAttribute('data-original-index', item.index);
                        errorElement.classList.add('translation-error');
                        translatedElements[item.index] = errorElement;
                    }
                });
                
                // 继续处理下一批次，而不是中断整个翻译过程
                console.log('继续处理下一批次...');
            }
        }
        
        // 清空加载提示
        translatedContent.innerHTML = '';
        
        // 按原始顺序添加翻译元素
        originalDocumentContent.forEach(item => {
            if (translatedElements[item.index]) {
                translatedContent.appendChild(translatedElements[item.index]);
            } else if (item.content.trim() !== '') {
                // 为没有翻译结果的元素创建一个占位符
                const placeholderElement = document.createElement(item.type);
                placeholderElement.textContent = `[未翻译: ${item.content}]`;
                placeholderElement.setAttribute('data-original-index', item.index);
                placeholderElement.classList.add('not-translated');
                translatedContent.appendChild(placeholderElement);
            }
        });
    }
    
    // 调用大模型API进行翻译
    async function translateTexts(texts, targetLanguage) {
        const modelType = modelTypeSelect.value;
        const languageMap = {
            'en': '英语',
            'zh': '中文',
            'ja': '日语',
            'ko': '韩语',
            'fr': '法语',
            'de': '德语',
            'es': '西班牙语',
            'ru': '俄语'
        };
        
        const targetLanguageName = languageMap[targetLanguage] || targetLanguage;
        
        if (modelType === 'openai') {
            const apiKey = document.getElementById('openai-api-key').value;
            const model = document.getElementById('openai-model').value;
            
            if (!apiKey) {
                throw new Error('请输入OpenAI API Key');
            }
            
            try {
                // 构建更结构化的提示词，明确指示翻译格式要求
                const prompt = `请将以下${texts.length}段文本翻译成${targetLanguageName}，保持原始格式和语义。

规则：
1. 只返回翻译结果，不要添加任何解释或额外内容
2. 保持原文的段落结构，每段翻译后用一个空行分隔
3. 保持原文的格式和标点符号风格
4. 确保翻译后的段落数量与原文相同

以下是需要翻译的文本，每段用空行分隔：

${texts.join('\n\n')}

翻译结果：`;
                
                const response = await fetch('https://api.openai.com/v1/chat/completions', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${apiKey}`
                    },
                    body: JSON.stringify({
                        model: model,
                        messages: [{ role: 'user', content: prompt }],
                        temperature: 0.3
                    })
                });
                
                const data = await response.json();
                
                if (!response.ok) {
                    throw new Error(data.error?.message || '翻译请求失败');
                }
                
                // 解析翻译结果
                console.log('OpenAI API 响应数据:', data);
                
                if (!data.choices || data.choices.length === 0 || !data.choices[0].message) {
                    console.error('OpenAI响应格式不正确:', data);
                    throw new Error('OpenAI响应格式不正确');
                }
                
                const translatedContent = data.choices[0].message.content.trim();
                console.log('提取的翻译内容:', translatedContent);
                
                // 分割翻译内容并确保返回正确数量的段落
                const translatedParagraphs = translatedContent.split('\n\n');
                console.log('分割后的段落数量:', translatedParagraphs.length, '原始文本数量:', texts.length);
                
                // 如果段落数量不匹配，尝试其他分割方法
                if (translatedParagraphs.length < texts.length) {
                    console.log('尝试使用其他分隔符分割翻译内容');
                    // 尝试使用单个换行符分割
                    const altSplit = translatedContent.split('\n').filter(line => line.trim() !== '');
                    if (altSplit.length >= texts.length) {
                        console.log('使用单个换行符分割成功');
                        return altSplit.slice(0, texts.length);
                    }
                }
                
                // 如果段落数量仍然不匹配，确保至少返回一些内容
                if (translatedParagraphs.length === 0) {
                    console.log('无法分割翻译内容，返回整个内容作为单个段落');
                    return [translatedContent];
                }
                
                return translatedParagraphs.slice(0, texts.length);
            } catch (error) {
                console.error('OpenAI 翻译错误:', error);
                throw error;
            }
        } else if (modelType === 'ollama') {
            const endpoint = document.getElementById('ollama-endpoint').value;
            const model = document.getElementById('ollama-model').value;
            
            if (!endpoint || !model) {
                throw new Error('请输入Ollama端点和模型名称');
            }
            
            try {
                // 构建更结构化的提示词，明确指示翻译格式要求
                const prompt = `请将以下${texts.length}段文本翻译成${targetLanguageName}，保持原始格式和语义。

规则：
1. 只返回翻译结果，不要添加任何解释或额外内容
2. 保持原文的段落结构，每段翻译后用一个空行分隔
3. 保持原文的格式和标点符号风格
4. 确保翻译后的段落数量与原文相同

以下是需要翻译的文本，每段用空行分隔：

${texts.join('\n\n')}

翻译结果：`;
                
                // 确保使用正确的API端点路径
                const apiUrl = `${endpoint}/api/generate`;
                console.log('Ollama API请求URL:', apiUrl);
                console.log('Ollama API请求参数:', {
                    model: model,
                    prompt: prompt,
                    stream: false
                });
                
                // 构建请求体，确保格式正确
                const requestBody = {
                    model: model,
                    prompt: prompt,
                    stream: false,
                    options: {
                        temperature: 0.3,  // 降低随机性，使翻译更准确
                        num_predict: 2048  // 确保有足够的输出长度
                    }
                };
                
                const response = await fetch(apiUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (!response.ok) {
                    throw new Error(data.error || '翻译请求失败');
                }
                
                // 解析翻译结果
                console.log('Ollama API 响应数据:', data);
                
                // 检查响应格式并提取翻译内容
                let translatedContent = '';
                if (data.response) {
                    translatedContent = data.response.trim();
                } else if (data.message) {
                    translatedContent = data.message.trim();
                } else if (data.content) {
                    translatedContent = data.content.trim();
                } else if (data.text) {
                    translatedContent = data.text.trim();
                } else {
                    console.error('无法从Ollama响应中提取翻译内容:', data);
                    throw new Error('无法从Ollama响应中提取翻译内容');
                }
                
                console.log('提取的翻译内容:', translatedContent);
                
                // 分割翻译内容并确保返回正确数量的段落
                const translatedParagraphs = translatedContent.split('\n\n');
                console.log('分割后的段落数量:', translatedParagraphs.length, '原始文本数量:', texts.length);
                
                // 如果段落数量不匹配，尝试其他分割方法
                if (translatedParagraphs.length < texts.length) {
                    console.log('尝试使用其他分隔符分割翻译内容');
                    // 尝试使用单个换行符分割
                    const altSplit = translatedContent.split('\n').filter(line => line.trim() !== '');
                    if (altSplit.length >= texts.length) {
                        console.log('使用单个换行符分割成功');
                        return altSplit.slice(0, texts.length);
                    }
                }
                
                // 如果段落数量仍然不匹配，确保至少返回一些内容
                if (translatedParagraphs.length === 0) {
                    console.log('无法分割翻译内容，返回整个内容作为单个段落');
                    return [translatedContent];
                }
                
                return translatedParagraphs.slice(0, texts.length);
            } catch (error) {
                console.error('Ollama 翻译错误:', error);
                throw error;
            }
        }
        
        throw new Error('不支持的模型类型');
    }
    
    // 保存为Word文档
    saveBtn.addEventListener('click', async () => {
        if (!currentFile || !originalDocumentContent.length === 0) {
            alert('没有可保存的文档');
            return;
        }
        
        try {
            // 获取翻译后的内容
            const translatedElements = translatedContent.querySelectorAll('[data-original-index]');
            const translatedTexts = {};
            
            translatedElements.forEach(element => {
                const originalIndex = element.getAttribute('data-original-index');
                translatedTexts[originalIndex] = element.textContent;
            });
            
            // 创建新的Word文档
            if (!window.docx) {
                throw new Error('docx库未正确加载，请刷新页面重试');
            }
            const doc = new window.docx.Document();
            
            // 根据原始文档结构和翻译内容创建新文档
            // 增强处理，保留更多格式信息
            const paragraphs = [];
            
            originalDocumentContent.forEach(item => {
                const translatedText = translatedTexts[item.index] || item.content;
                
                // 提取样式信息
                const styleInfo = item.styleInfo || {};
                const styledChildren = item.styledChildren || [];
                
                // 创建文本运行对象，用于设置格式
                let textRuns = [];
                
                // 如果有带样式的子元素，为每个子元素创建单独的文本运行
                if (styledChildren.length > 0) {
                    // 这里简化处理，实际情况可能需要更复杂的文本匹配算法
                    styledChildren.forEach(child => {
                        // 从样式类名中提取颜色和格式信息
                        const isRed = child.classList.includes('red');
                        const isBlue = child.classList.includes('blue');
                        const isBold = child.classList.includes('bold');
                        const isItalic = child.classList.includes('italic');
                        const isUnderline = child.classList.includes('underline');
                        
                        // 提取自定义颜色
                        let customColor = null;
                        child.classList.forEach(className => {
                            if (className.startsWith('color-')) {
                                customColor = className.substring(6); // 去掉'color-'前缀
                            }
                        });
                        
                        // 提取字体
                        let customFont = null;
                        child.classList.forEach(className => {
                            if (className.startsWith('font-')) {
                                customFont = className.substring(5).replace(/-/g, ' '); // 去掉'font-'前缀并还原空格
                            }
                        });
                        
                        // 提取字号
                        let customSize = null;
                        child.classList.forEach(className => {
                            if (className.startsWith('size-')) {
                                customSize = parseInt(className.substring(5)); // 去掉'size-'前缀并转为数字
                            }
                        });
                        
                        // 从样式对象中提取颜色和字体信息
                        const childStyles = child.styles || {};
                        
                        // 创建带格式的文本运行
                        textRuns.push(new window.docx.TextRun({
                            text: child.text,
                            bold: isBold || childStyles.fontWeight === 'bold' || childStyles.fontWeight >= 700,
                            italics: isItalic || childStyles.fontStyle === 'italic',
                            underline: isUnderline || childStyles.textDecoration.includes('underline'),
                            color: customColor || (isRed ? 'FF0000' : (isBlue ? '0000FF' : childStyles.color.replace(/[^0-9A-Fa-f]/g, ''))),
                            size: customSize || parseInt(childStyles.fontSize) * 2, // 转换为半点值
                            font: customFont || childStyles.fontFamily.split(',')[0].replace(/['"\/]/g, '')
                        }));
                    });
                } else {
                    // 没有特殊格式的子元素，创建单一文本运行
                    textRuns.push(new window.docx.TextRun({
                        text: translatedText,
                        bold: styleInfo.fontWeight === 'bold' || styleInfo.fontWeight >= 700,
                        italics: styleInfo.fontStyle === 'italic',
                        underline: styleInfo.textDecoration && styleInfo.textDecoration.includes('underline'),
                        color: styleInfo.color ? styleInfo.color.replace(/[^0-9A-Fa-f]/g, '') : undefined,
                        size: styleInfo.fontSize ? parseInt(styleInfo.fontSize) * 2 : undefined, // 转换为半点值
                        font: styleInfo.fontFamily ? styleInfo.fontFamily.split(',')[0].replace(/['"\/]/g, '') : undefined
                    }));
                }
                
                // 根据元素类型创建段落
                let paragraph;
                switch (item.type) {
                    case 'h1':
                        paragraph = new window.docx.Paragraph({
                            children: textRuns,
                            heading: window.docx.HeadingLevel.HEADING_1
                        });
                        break;
                    case 'h2':
                        paragraph = new window.docx.Paragraph({
                            children: textRuns,
                            heading: window.docx.HeadingLevel.HEADING_2
                        });
                        break;
                    case 'h3':
                        paragraph = new window.docx.Paragraph({
                            children: textRuns,
                            heading: window.docx.HeadingLevel.HEADING_3
                        });
                        break;
                    case 'li':
                        paragraph = new window.docx.Paragraph({
                            children: textRuns,
                            bullet: {
                                level: 0
                            }
                        });
                        break;
                    default:
                        paragraph = new window.docx.Paragraph({
                            children: textRuns
                        });
                }
                
                paragraphs.push(paragraph);
            });
            
            doc.addSection({
                properties: {},
                children: paragraphs
            });
            
            // 生成文档并保存
            const buffer = await window.docx.Packer.toBlob(doc);
            const fileName = currentFile.name.replace('.docx', '_translated.docx');
            saveAs(buffer, fileName);
        } catch (error) {
            console.error('保存文档错误:', error);
            alert(`保存文档错误: ${error.message}`);
        }
    });
});