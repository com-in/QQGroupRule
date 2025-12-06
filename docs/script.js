// 全局变量
let canvas, ctx;
let isDrawing = false;
let currentTool = 'brush'; // 默认工具：画笔
let currentColor = '#ff0000'; // 默认颜色：红色
let lastX = 0;
let lastY = 0;
let arrowStartX = 0;
let arrowStartY = 0;
let isDrawingArrow = false;
let isSlideShowMode = false;
let isOffice = false;
let isWPS = false;

// 初始化函数
function init() {
    // 获取画布元素
    canvas = document.getElementById('canvas');
    ctx = canvas.getContext('2d');
    
    // 设置画布大小
    resizeCanvas();
    
    // 添加窗口大小改变事件
    window.addEventListener('resize', resizeCanvas);
    
    // 初始化事件监听
    initEventListeners();
    
    // 初始化RGB调色盘
    initRGBPicker();
    
    // 设置初始颜色预览
    updateColorPreview();
    
    // 初始化Office/WPS集成
    initOfficeWPSIntegration();
}

// 调整画布大小
function resizeCanvas() {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
}

// 初始化事件监听
function initEventListeners() {
    // 工具栏事件
    document.getElementById('arrow-btn').addEventListener('click', () => selectTool('arrow'));
    document.getElementById('brush-btn').addEventListener('click', () => selectTool('brush'));
    document.getElementById('eraser-btn').addEventListener('click', () => selectTool('eraser'));
    
    // 颜色选择事件
    const colorBtns = document.querySelectorAll('.color-btn');
    colorBtns.forEach(btn => {
        btn.addEventListener('click', () => selectColor(btn.dataset.color));
    });
    
    // 更多菜单事件
    document.getElementById('more-btn').addEventListener('click', toggleMoreMenu);
    document.getElementById('close-dropdown').addEventListener('click', closeMoreMenu);
    
    // 点击页面其他地方关闭更多菜单
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.more-menu')) {
            closeMoreMenu();
        }
    });
    
    // 画布事件
    canvas.addEventListener('mousedown', startDrawing);
    canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDrawing);
    canvas.addEventListener('mouseout', stopDrawing);
    
    // 触摸事件支持
    canvas.addEventListener('touchstart', (e) => {
        e.preventDefault();
        const touch = e.touches[0];
        const mouseEvent = new MouseEvent('mousedown', {
            clientX: touch.clientX,
            clientY: touch.clientY
        });
        canvas.dispatchEvent(mouseEvent);
    });
    
    canvas.addEventListener('touchmove', (e) => {
        e.preventDefault();
        const touch = e.touches[0];
        const mouseEvent = new MouseEvent('mousemove', {
            clientX: touch.clientX,
            clientY: touch.clientY
        });
        canvas.dispatchEvent(mouseEvent);
    });
    
    canvas.addEventListener('touchend', (e) => {
        e.preventDefault();
        const mouseEvent = new MouseEvent('mouseup', {});
        canvas.dispatchEvent(mouseEvent);
    });
}

// 初始化RGB调色盘
function initRGBPicker() {
    const redSlider = document.getElementById('red-slider');
    const greenSlider = document.getElementById('green-slider');
    const blueSlider = document.getElementById('blue-slider');
    
    const redValue = document.getElementById('red-value');
    const greenValue = document.getElementById('green-value');
    const blueValue = document.getElementById('blue-value');
    
    // 添加滑块事件
    redSlider.addEventListener('input', updateRGBColor);
    greenSlider.addEventListener('input', updateRGBColor);
    blueSlider.addEventListener('input', updateRGBColor);
    
    // 更新RGB值显示
    function updateRGBColor() {
        redValue.textContent = redSlider.value;
        greenValue.textContent = greenSlider.value;
        blueValue.textContent = blueSlider.value;
        
        // 更新颜色预览
        updateColorPreview();
        
        // 自动选择自定义颜色
        const customColor = rgbToHex(redSlider.value, greenSlider.value, blueSlider.value);
        selectColor(customColor, true);
    }
}

// 更新颜色预览
function updateColorPreview() {
    const redSlider = document.getElementById('red-slider');
    const greenSlider = document.getElementById('green-slider');
    const blueSlider = document.getElementById('blue-slider');
    
    const colorPreview = document.getElementById('color-preview');
    colorPreview.style.backgroundColor = `rgb(${redSlider.value}, ${greenSlider.value}, ${blueSlider.value})`;
}

// 选择工具
function selectTool(tool) {
    // 更新当前工具
    currentTool = tool;
    
    // 移除所有工具按钮的活跃状态
    const toolBtns = document.querySelectorAll('.tool-btn');
    toolBtns.forEach(btn => btn.classList.remove('active'));
    
    // 激活当前工具按钮
    document.getElementById(`${tool}-btn`).classList.add('active');
    
    // 更新鼠标指针样式
    updateCursorStyle();
}

// 更新鼠标指针样式
function updateCursorStyle() {
    switch(currentTool) {
        case 'arrow':
            canvas.style.cursor = 'crosshair';
            break;
        case 'brush':
            canvas.style.cursor = 'crosshair';
            break;
        case 'eraser':
            canvas.style.cursor = 'cell';
            break;
    }
}

// 选择颜色
function selectColor(color, isCustom = false) {
    // 更新当前颜色
    currentColor = color;
    
    // 移除所有颜色按钮的活跃状态
    const colorBtns = document.querySelectorAll('.color-btn');
    colorBtns.forEach(btn => btn.classList.remove('active'));
    
    // 如果不是自定义颜色，激活对应的颜色按钮
    if (!isCustom) {
        const selectedBtn = document.querySelector(`[data-color="${color}"]`);
        if (selectedBtn) {
            selectedBtn.classList.add('active');
        }
        
        // 更新RGB滑块值
        const rgb = hexToRgb(color);
        document.getElementById('red-slider').value = rgb.r;
        document.getElementById('green-slider').value = rgb.g;
        document.getElementById('blue-slider').value = rgb.b;
        
        // 更新RGB值显示
        document.getElementById('red-value').textContent = rgb.r;
        document.getElementById('green-value').textContent = rgb.g;
        document.getElementById('blue-value').textContent = rgb.b;
        
        // 更新颜色预览
        updateColorPreview();
    }
}

// 开始绘制
function startDrawing(e) {
    isDrawing = true;
    
    // 获取鼠标位置
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    // 根据当前工具执行不同的操作
    switch(currentTool) {
        case 'arrow':
            isDrawingArrow = true;
            arrowStartX = x;
            arrowStartY = y;
            break;
        case 'brush':
            lastX = x;
            lastY = y;
            // 设置画笔样式
            ctx.strokeStyle = currentColor;
            ctx.lineWidth = 3;
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
            break;
        case 'eraser':
            lastX = x;
            lastY = y;
            // 设置橡皮擦样式
            ctx.strokeStyle = 'rgba(0, 0, 0, 0)';
            ctx.lineWidth = 20;
            ctx.lineCap = 'round';
            ctx.globalCompositeOperation = 'destination-out';
            break;
    }
}

// 绘制
function draw(e) {
    if (!isDrawing) return;
    
    // 获取鼠标位置
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    // 根据当前工具执行不同的绘制操作
    switch(currentTool) {
        case 'arrow':
            if (isDrawingArrow) {
                // 清除画布上的临时箭头
                clearTempArrow();
                // 绘制临时箭头
                drawTempArrow(arrowStartX, arrowStartY, x, y);
            }
            break;
        case 'brush':
            // 绘制自由线条
            ctx.beginPath();
            ctx.moveTo(lastX, lastY);
            ctx.lineTo(x, y);
            ctx.stroke();
            lastX = x;
            lastY = y;
            break;
        case 'eraser':
            // 擦除操作
            ctx.beginPath();
            ctx.moveTo(lastX, lastY);
            ctx.lineTo(x, y);
            ctx.stroke();
            lastX = x;
            lastY = y;
            break;
    }
}

// 停止绘制
function stopDrawing(e) {
    if (!isDrawing) return;
    
    isDrawing = false;
    
    // 获取鼠标位置
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    // 根据当前工具执行不同的停止操作
    switch(currentTool) {
        case 'arrow':
            if (isDrawingArrow) {
                // 绘制最终箭头
                drawArrow(arrowStartX, arrowStartY, x, y);
                isDrawingArrow = false;
            }
            break;
        case 'eraser':
            // 恢复正常绘制模式
            ctx.globalCompositeOperation = 'source-over';
            break;
    }
}

// 绘制箭头
function drawArrow(x1, y1, x2, y2) {
    const arrowSize = 15;
    const angle = Math.atan2(y2 - y1, x2 - x1);
    
    // 设置箭头样式
    ctx.strokeStyle = currentColor;
    ctx.fillStyle = currentColor;
    ctx.lineWidth = 3;
    
    // 绘制直线
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    
    // 绘制箭头头部
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(
        x2 - arrowSize * Math.cos(angle - Math.PI / 6),
        y2 - arrowSize * Math.sin(angle - Math.PI / 6)
    );
    ctx.lineTo(
        x2 - arrowSize * Math.cos(angle + Math.PI / 6),
        y2 - arrowSize * Math.sin(angle + Math.PI / 6)
    );
    ctx.closePath();
    ctx.fill();
}

// 绘制临时箭头（用于拖拽过程中）
function drawTempArrow(x1, y1, x2, y2) {
    const arrowSize = 15;
    const angle = Math.atan2(y2 - y1, x2 - x1);
    
    // 保存当前状态
    ctx.save();
    
    // 设置临时箭头样式（半透明）
    ctx.strokeStyle = currentColor;
    ctx.fillStyle = currentColor;
    ctx.lineWidth = 3;
    ctx.globalAlpha = 0.7;
    
    // 绘制直线
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    
    // 绘制箭头头部
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(
        x2 - arrowSize * Math.cos(angle - Math.PI / 6),
        y2 - arrowSize * Math.sin(angle - Math.PI / 6)
    );
    ctx.lineTo(
        x2 - arrowSize * Math.cos(angle + Math.PI / 6),
        y2 - arrowSize * Math.sin(angle + Math.PI / 6)
    );
    ctx.closePath();
    ctx.fill();
    
    // 恢复状态
    ctx.restore();
}

// 清除临时箭头
function clearTempArrow() {
    // 简单实现：清除整个画布并重绘
    // 实际项目中应该使用离屏画布或保存绘制历史
    ctx.clearRect(0, 0, canvas.width, canvas.height);
}

// 切换更多菜单
function toggleMoreMenu() {
    const dropdown = document.getElementById('dropdown-menu');
    dropdown.classList.toggle('show');
}

// 关闭更多菜单
function closeMoreMenu() {
    const dropdown = document.getElementById('dropdown-menu');
    dropdown.classList.remove('show');
}

// RGB转十六进制颜色
function rgbToHex(r, g, b) {
    return '#' + ((1 << 24) + (parseInt(r) << 16) + (parseInt(g) << 8) + parseInt(b)).toString(16).slice(1);
}

// 十六进制颜色转RGB
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}

// 初始化Office/WPS集成
function initOfficeWPSIntegration() {
    // 检测环境
    isOffice = typeof Office !== 'undefined';
    isWPS = typeof WPS !== 'undefined';
    
    if (isOffice) {
        // Office环境初始化
        Office.onReady((info) => {
            if (info.host === Office.HostType.PowerPoint) {
                console.log('PPT绘图工具已加载到Office PowerPoint中');
                setupOfficeEventListeners();
            } else {
                console.error('此插件仅适用于PowerPoint');
                hidePlugin();
            }
        });
    } else if (isWPS) {
        // WPS环境初始化
        WPS.ready(() => {
            console.log('PPT绘图工具已加载到WPS演示中');
            setupWPSEventListeners();
        });
    } else {
        // 非PPT环境，隐藏插件
        console.error('此插件仅适用于Office PowerPoint或WPS演示');
        hidePlugin();
    }
}

// 隐藏插件（非PPT环境下）
function hidePlugin() {
    // 隐藏工具栏和画布
    const toolbar = document.getElementById('toolbar');
    toolbar.classList.add('hidden');
    canvas.style.display = 'none';
    
    // 显示错误信息
    const errorMsg = document.createElement('div');
    errorMsg.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background-color: rgba(0, 0, 0, 0.8);
        color: white;
        padding: 20px;
        border-radius: 10px;
        font-size: 16px;
        text-align: center;
        z-index: 10001;
    `;
    errorMsg.textContent = '此插件仅适用于Office PowerPoint或WPS演示，无法作为独立网页运行。';
    document.body.appendChild(errorMsg);
}

// 设置Office事件监听器
function setupOfficeEventListeners() {
    try {
        // 检测是否在幻灯片放映模式
        if (Office.context.requirements.isSetSupported('PowerPointApi', '1.5')) {
            // 添加幻灯片放映事件监听
            Office.context.document.addHandlerAsync(
                Office.EventType.PowerPointSlideShowBegin,
                onSlideShowBegin
            );
            
            Office.context.document.addHandlerAsync(
                Office.EventType.PowerPointSlideShowEnd,
                onSlideShowEnd
            );
            
            // 检查当前是否已经在幻灯片放映模式
            checkCurrentSlideShowState();
        } else {
            console.log('当前Office版本不支持幻灯片放映事件');
            // 降级处理：始终显示工具栏
            isSlideShowMode = true;
            showToolbar();
        }
    } catch (error) {
        console.error('设置Office事件监听失败:', error);
        // 降级处理：始终显示工具栏
        isSlideShowMode = true;
        showToolbar();
    }
}

// 设置WPS事件监听器
function setupWPSEventListeners() {
    try {
        // WPS演示幻灯片放映事件
        if (WPS.ApiVersion >= 1000) {
            // 添加幻灯片放映开始事件
            WPS.Events.Subscribe("SlideShowBegin", onSlideShowBegin);
            // 添加幻灯片放映结束事件
            WPS.Events.Subscribe("SlideShowEnd", onSlideShowEnd);
            
            // 检查当前是否已经在幻灯片放映模式
            checkCurrentWPSState();
        } else {
            console.log('当前WPS版本不支持幻灯片放映事件');
            // 降级处理：始终显示工具栏
            isSlideShowMode = true;
            showToolbar();
        }
    } catch (error) {
        console.error('设置WPS事件监听失败:', error);
        // 降级处理：始终显示工具栏
        isSlideShowMode = true;
        showToolbar();
    }
}

// 检查当前Office幻灯片放映状态
function checkCurrentSlideShowState() {
    try {
        // 使用Office API检查当前状态
        // 注意：Office.js可能没有直接的API来检查幻灯片放映状态
        // 这里使用降级方案，假设插件只在幻灯片放映时加载
        isSlideShowMode = true;
        showToolbar();
    } catch (error) {
        console.error('检查幻灯片放映状态失败:', error);
        isSlideShowMode = true;
        showToolbar();
    }
}

// 检查当前WPS幻灯片放映状态
function checkCurrentWPSState() {
    try {
        // 使用WPS API检查当前状态
        const app = WPS.Application;
        isSlideShowMode = app.SlideShowWindows.Count > 0;
        if (isSlideShowMode) {
            showToolbar();
        } else {
            hideToolbar();
        }
    } catch (error) {
        console.error('检查WPS幻灯片放映状态失败:', error);
        isSlideShowMode = true;
        showToolbar();
    }
}

// 幻灯片放映开始事件处理
function onSlideShowBegin() {
    console.log('幻灯片放映已开始');
    isSlideShowMode = true;
    showToolbar();
}

// 幻灯片放映结束事件处理
function onSlideShowEnd() {
    console.log('幻灯片放映已结束');
    isSlideShowMode = false;
    hideToolbar();
    // 清除画布内容
    ctx.clearRect(0, 0, canvas.width, canvas.height);
}

// 显示工具栏
function showToolbar() {
    const toolbar = document.getElementById('toolbar');
    toolbar.classList.remove('hidden');
    canvas.style.pointerEvents = 'auto';
}

// 隐藏工具栏
function hideToolbar() {
    const toolbar = document.getElementById('toolbar');
    toolbar.classList.add('hidden');
    canvas.style.pointerEvents = 'none';
}

// 页面加载完成后初始化
window.addEventListener('DOMContentLoaded', init);

// 监听来自父窗口的消息
window.addEventListener('message', (event) => {
    if (event.data.action === 'clearCanvas') {
        // 清除画布内容
        ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
});