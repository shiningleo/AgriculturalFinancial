$(function () { 
    const container = $('.combox-wrapper');
    const input = container.find('.combox-input');
    const listContainer = container.find('.combox-list-container');
    const listItems = ['apple', 'banana', 'cherry', 'durian', 'grape', 'kiwi', 'lemon', 'mango', 'orange', 'peach', 'pear', 'watermelon'];

    function showList() {
        // 显示下拉菜单
        listContainer.css('display', 'block');

        // 清除列表项
        listContainer.html('');

        // 输入框的值
        const inputValue = input.val();

        // 根据输入的值过滤选项
        const filteredList = listItems.filter((item) => {
            return item.includes(inputValue);
        });

        // 将过滤后的选项填充到下拉菜单中
        filteredList.forEach((item) => {
            const li = $('<li>').text(item);
            li.off('click').on('click', function () { // 修改此处
                input.val($(this).text());
                listContainer.css('display', 'none');
            });
            listContainer.append(li);
        });
    }
    showList();
   // console.log($('.combox-container')); // 确认输出的是一个 jQuery 对象
    //$("#richbox1 .richbox-item").draggable({
    //    helper: "clone",
    //    revert: "invalid"
    //});

    //$("#richbox1").droppable({
    //    drop: function (event, ui) {
    //        const droppable = $(this);
    //        const draggable = ui.draggable;
    //        const newElement = draggable.clone().appendTo(droppable);

    //        // 变更位置
    //        draggable.css({
    //            top: 0,
    //            left: 0
    //        });
    //        newElement.css({
    //            position: "relative",
    //            top: 0,
    //            left: 0
    //        });

    //        // 删除原拖拽元素
    //        draggable.remove();
    //    }
    //});


    //拖动每一项目
    // 为每个richbox添加拖放功能
    //$('textarea[id^="richbox"]').draggable({
    //    revert: true, // 拖放结束后richbox返回原位置
    //    helper: 'clone', // 使用克隆来拖放
    //});

    //$('textarea[id^="richbox"]').droppable({
    //    accept: function (draggable) {
    //        // richbox6只能被拖入到自己的位置
    //        if ($(this).attr('id') === 'richbox6' && $(draggable).attr('id') !== 'richbox6') {
    //            return false;
    //        }
    //        return true;
    //    },
    //    drop: function (event, ui) {
    //        // 将拖动项添加到目标richbox中
    //        $(this).val(function (index, value) {
    //            return value + '\n' + ui.draggable.text().trim();
    //        });
    //    }
    //});

    //// 给每个richbox的行添加排序功能
    //$('.form-control').sortable({
    //    handle: '.sortable-handle', // 拖动把手所在的元素
    //    connectWith: '.form-control', // 允许和其它richbox连接
    //});

    //function run() {
    //    Excel.run( (context) => {
    //        const range = context.workbook.getSelectedRange();
    //        range.format.fill.color = "yellow";
    //        range.load("address");

    //       // context.sync();

    //        console.log(`The range address was "${range.address}".`);
    //    });
    //}
    $(".context-menu").on("contextmenu", function (e) {
        // 阻止默认的浏览器右键菜单弹出
        e.preventDefault();

        // 获取选择的文本行
        var textarea = e.currentTarget;
       // var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
        textarea.addEventListener("click", handleClick);
        // 定义右键菜单的项和回调函数
        var menuItems = {
            moveUp: {
                name: "上移",
                icon: "fas fa-arrow-up",
                // disabled: selectedLine.length === 0
            },
            moveDown: {
                name: "下移",
                icon: "fas fa-arrow-down",
                //disabled: selectedLine.length === 0
            },
            moveToTop: {
                name: "移到开头",
                icon: "fas fa-angle-double-up",
                // disabled: selectedLine.length === 0
            },
            moveToBottom: {
                name: "移到末尾",
                icon: "fas fa-angle-double-down",
                // disabled: selectedLine.length === 0
            },
            moveTo: {
                name: "移到其他框体",
                icon: "fas fa-exchange-alt",
                //disabled: selectedLine.length === 0,
                items: {
                    richbox1: {
                        name: "移动到报表筛选",
                        disabled: textarea.id === "richbox1",
                        callback: function () {
                            moveTo(textarea.id, "richbox1");
                        }
                    },
                    richbox2: {
                        name: "移动到列标签",
                        disabled: textarea.id === "richbox2",
                        callback: function () {
                            moveTo(textarea.id, "richbox2");
                        }
                    },
                    richbox3: {
                        name: "移动到行标签",
                        disabled: textarea.id === "richbox3",
                        callback: function () {
                            moveTo(textarea.id, "richbox3");
                        }
                    },
                    richbox4: {
                        name: "移动到值标签",
                        disabled: textarea.id === "richbox4",
                        callback: function () {
                            moveTo(textarea.id, "richbox4");
                        }
                    },
                    richbox5: {
                        name: "移动到指标标签",
                        disabled: textarea.id === "richbox5",
                        callback: function () {
                            moveTo(textarea.id, "richbox5");
                        }
                    },
                    richbox6: {
                        name: "移动到维度标签",
                        disabled: textarea.id === "richbox6",
                        callback: function () {
                            moveTo(textarea.id, "richbox6");
                        }
                    }
                    // 其他 richbox

                }
            },
            remove: {
                name: "删除",
                icon: "fas fa-trash",
                // disabled: selectedLine.length === 0
            }
        }
        // 调用 jQuery ContextMenu 插件来创建菜单
        $.contextMenu({
            selector: "#" + textarea.id,
            callback: function (key) {
                // 根据菜单的名称执行相应的操作
                switch (key) {
                    case "moveUp":
                        moveUp(textarea.id);
                        break;
                    case "moveDown":
                        moveDown(textarea.id);
                        break;
                    case "moveToTop":
                        moveToTop(textarea.id);
                        break;
                    case "moveToBottom":
                        moveToBottom(textarea.id);
                        break;
                    case "remove":
                        remove(textarea.id);
                        break;
                    // 其他 richbox
                }
                // 在此处添加调用 moveToTop 函数的代码
                if (key === "moveToTop") {
                    moveToTop(textarea.id);
                }
                // 在此处添加调用 moveToTop 函数的代码
            },
            items: menuItems
        })

    })

    //下拉cube
    selectcube();
 


});

function selectcube() {

    // 定义下拉菜单选项数组
    var options = ['cube1', 'cube2', 'cube3'];

    // 动态添加下拉菜单选项
    for (var i = 0; i < options.length; i++) {
        $('#selectcube').append('<option value="' + options[i] + '">' + options[i] + '</option>');
    }
}  

class TabsAutomatic {
    constructor(groupNode) {
        this.tablistNode = groupNode;

        this.tabs = [];

        this.firstTab = null;
        this.lastTab = null;

        this.tabs = Array.from(this.tablistNode.querySelectorAll('[role=tab]'));
        this.tabpanels = [];

        for (var i = 0; i < this.tabs.length; i += 1) {
            var tab = this.tabs[i];
            var tabpanel = document.getElementById(tab.getAttribute('aria-controls'));

            tab.tabIndex = -1;
            tab.setAttribute('aria-selected', 'false');
            this.tabpanels.push(tabpanel);

            tab.addEventListener('keydown', this.onKeydown.bind(this));
            tab.addEventListener('click', this.onClick.bind(this));

            if (!this.firstTab) {
                this.firstTab = tab;
            }
            this.lastTab = tab;
        }

        this.setSelectedTab(this.firstTab, false);
    }

    setSelectedTab(currentTab, setFocus) {
        if (typeof setFocus !== 'boolean') {
            setFocus = true;
        }
        for (var i = 0; i < this.tabs.length; i += 1) {
            var tab = this.tabs[i];
            if (currentTab === tab) {
                tab.setAttribute('aria-selected', 'true');
                tab.removeAttribute('tabindex');
                this.tabpanels[i].classList.remove('is-hidden');
                if (setFocus) {
                    tab.focus();
                }
            } else {
                tab.setAttribute('aria-selected', 'false');
                tab.tabIndex = -1;
                this.tabpanels[i].classList.add('is-hidden');
            }
        }
    }

    setSelectedToPreviousTab(currentTab) {
        var index;

        if (currentTab === this.firstTab) {
            this.setSelectedTab(this.lastTab);
        } else {
            index = this.tabs.indexOf(currentTab);
            this.setSelectedTab(this.tabs[index - 1]);
        }
    }

    setSelectedToNextTab(currentTab) {
        var index;

        if (currentTab === this.lastTab) {
            this.setSelectedTab(this.firstTab);
        } else {
            index = this.tabs.indexOf(currentTab);
            this.setSelectedTab(this.tabs[index + 1]);
        }
    }

    /* EVENT HANDLERS */

    onKeydown(event) {
        var tgt = event.currentTarget,
            flag = false;

        switch (event.key) {
            case 'ArrowLeft':
                this.setSelectedToPreviousTab(tgt);
                flag = true;
                break;

            case 'ArrowRight':
                this.setSelectedToNextTab(tgt);
                flag = true;
                break;

            case 'Home':
                this.setSelectedTab(this.firstTab);
                flag = true;
                break;

            case 'End':
                this.setSelectedTab(this.lastTab);
                flag = true;
                break;

            default:
                break;
        }

        if (flag) {
            event.stopPropagation();
            event.preventDefault();
        }
    }

    onClick(event) {
        this.setSelectedTab(event.currentTarget);
    }
}

// Initialize tablist

window.addEventListener('load', function () {
    var tablists = document.querySelectorAll('[role=tablist].automatic');
    for (var i = 0; i < tablists.length; i++) {
        new TabsAutomatic(tablists[i]);
    }
});


//const chatForm = document.querySelector('.chat-form');
//const chatInput = document.querySelector('.chat-input');
//const chatMessages = document.querySelector('.chat-messages');
//const chatBtnSend = document.querySelector('.chat-btn-send');
//const chatBtnSetting = document.querySelector('.chat-btn-setting');
//// 初始化富文本编辑器
//const quill = new Quill(chatInput, {
//    theme: 'snow'
//});


//// 聊天发送按钮的点击事件
//chatBtnSend.addEventListener('click', sendMessage);

//function sendMessage(e) {

//    e.preventDefault(); // 阻止表单默认提交行为

//    // 获取输入的消息文本
//    const message = quill.root.innerHTML.trim();

//    // 如果消息为空，直接返回
//    if (!message) return;

//    // 创建一个新的聊天消息
//    const chatMessage = document.createElement('div');
//    chatMessage.classList.add('chat-message');
//    chatMessage.innerHTML = message;

//    // 将新的聊天消息添加到聊天列表中
//    chatMessages.appendChild(chatMessage);

//    // 清空输入框
//    quill.root.innerHTML = '';
//    // 自动滚动到底部
//    chatMessages.scrollTop = chatMessages.scrollHeight;
//}
//上移
function moveUp(id) {
    var textarea = document.getElementById(id);
   // var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var selectedLine = handleClick(id);
    // if (selectedLine.length > 0) {

    if (selectedLine.trim() === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    var lines = textarea.value.trim().split("\n").map(l => l.trim());
    var lineIndex = lines.indexOf(selectedLine.trim());

    if (lineIndex > 0) {
        lines.splice(lineIndex, 1);
        lines.splice(lineIndex - 1, 0, selectedLine);
        textarea.value = lines.join("\n");
    }
    //}
}
//下移
function moveDown(id) {
    var textarea = document.getElementById(id);
    //var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var selectedLine = handleClick(id);
    //  if (selectedLine.length > 0) {
    if (selectedLine.trim() === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    var lines = textarea.value.trim().split("\n").map(l => l.trim());
    var lineIndex = lines.indexOf(selectedLine.trim());

    if (lineIndex < lines.length - 1) {
        lines.splice(lineIndex, 1);
        lines.splice(lineIndex + 1, 0, selectedLine);
        textarea.value = lines.join("\n");
    }
    // }
}
//移动
function moveTo(id1, id2) {
    var fromId = id1;
    var toId = id2;
    var textarea1 = document.getElementById(id1);
    var textarea2 = document.getElementById(id2);
    var selectedLine = handleClick(id1);
    if (selectedLine === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    // 判断是否可以移动
    if (fromId === "richbox4" || (fromId === "richbox5" && toId !== "richbox4")) {
        return; // 如果源文本框为 richbox4 或 richbox5 但目标文本框不是 richbox5，则不能移动
    }
    // 判断是否可以移动
    if (fromId === "richbox6"&&toId=="richbox4") {
        return; // 如果源文本框为 richbox6 但目标文本框是 richbox4，则不能移动
    }
    // 判断是否可以移动
    if (fromId === "richbox1" || fromId === "richbox2" || fromId === "richbox3" && toId == "richbox4") {
        return; // 如果源文本框为 richbox6 但目标文本框是 richbox4，则不能移动
    }
    // 判断是否可以移动到指定的文本框中
    if ((toId === "richbox1" || toId === "richbox2" || toId === "richbox3" || toId === "richbox4" || toId === "richbox6") && fromId !== toId) {
        var lines1 = textarea1.value.trim().split("\n").map(l => l.trim());
        var lines2 = textarea2.value.trim().split("\n").map(l => l.trim());
        var lineIndex = lines1.indexOf(selectedLine.trim());

        if (lineIndex >= 0) {
            lines1.splice(lineIndex, 1);
            lines2.push(selectedLine);
            textarea1.value = lines1.join("\n");
            textarea2.value = lines2.join("\n");
        }
        //}
    }
}
//移除
function remove(id) {
    var textarea = document.getElementById(id);
    //var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var selectedLine = handleClick(id);
    // if (selectedLine.length > 0) {
    if (selectedLine.trim() === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    var lines = textarea.value.trim().split("\n").map(l => l.trim());
    var lineIndex = lines.indexOf(selectedLine.trim());

    if (lineIndex >= 0) {
        lines.splice(lineIndex, 1);
        textarea.value = lines.join("\n");
    }
    //}
}
//移到开头
function moveToTop(id) {
    var textarea = document.getElementById(id);
   // var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var selectedLine = handleClick(id);
    //  if (selectedLine.length > 0) {
    if (selectedLine.trim() === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    var lines = textarea.value.trim().split("\n").map(l => l.trim());
    var lineIndex = lines.indexOf(selectedLine.trim());

    if (lineIndex > 0) {
        lines.splice(lineIndex, 1);
        lines.splice(0, 0, selectedLine);
        textarea.value = lines.join("\n");
    }
    //}
}
//移到末尾
function moveToBottom(id) {
    var textarea = document.getElementById(id);
    //var selectedLine = textarea.value.substring(textarea.selectionStart, textarea.selectionEnd);
    var selectedLine = handleClick(id);
    // if (selectedLine.length > 0) {
    if (selectedLine.trim() === "") {
        return; // 如果没有选中整行，不执行任何操作
    }
    var lines = textarea.value.trim().split("\n").map(l => l.trim());
    var lineIndex = lines.indexOf(selectedLine.trim());

    if (lineIndex >= 0 && lineIndex < lines.length - 1) {
        lines.splice(lineIndex, 1);
        lines.push(selectedLine);
        textarea.value = lines.join("\n");
    }
    //}
}


// 鼠标点击文本区域时触发的函数 
function handleClick(id) {
    var textarea = document.getElementById(id);
    var cursorPos = textarea.selectionStart;
    var text = textarea.value;
    var lineStart = text.lastIndexOf("\n", cursorPos - 1) + 1;
    var lineEnd = text.indexOf("\n", cursorPos);
    if (lineEnd == -1) {
        lineEnd = text.length;
    }
    var selectedLine = text.substring(lineStart, lineEnd);
    return selectedLine;
    console.log(selectedLine);
}
 




