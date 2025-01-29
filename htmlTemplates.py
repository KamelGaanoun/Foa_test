upload_style="""
<style>
/* Customize the drag-and-drop instructions */
[data-testid='stFileUploaderDropzoneInstructions'] > div > span {
    display: none;
}
[data-testid='stFileUploaderDropzoneInstructions'] > div::before {
    content: 'Chargez ou glissez votre fichier ici';
    visibility: visible;
    display: block;
}

/* Target only the "Browse" button of the file uploader */
[data-testid='stFileUploader'] button[data-testid='stBaseButton-secondary'] {
    text-indent: -9999px;
    line-height: 0;
}
[data-testid='stFileUploader'] button[data-testid='stBaseButton-secondary']::after {
    line-height: initial;
    content: "Parcourir";
    text-indent: 0;
    visibility: visible;
    display: block;
}

/* Target the file size limit */
[data-testid='stFileUploader'] small {
    visibility: hidden;
}
[data-testid='stFileUploader'] small::after {
    content: "Limite 200MB par fichier";
    visibility: visible;
    display: block;
}
</style>
"""

toggle_switch="""
<style>
    .toggle { 
        -webkit-appearance: none; 
        appearance: none;
        background-color: #4CAF50; 
        width: 60px; 
        height: 30px; 
        border-radius: 50px; 
        position: relative;
    }
    .toggle:checked { 
        background-color: #2196F3; 
    }
    .toggle:before {
        content: ''; 
        position: absolute; 
        top: 3px; 
        left: 3px; 
        width: 24px; 
        height: 24px; 
        border-radius: 50%; 
        background-color: white; 
        transition: 0.3s;
    }
    .toggle:checked:before {
        left: 33px;
    }
</style>
<label for="save-toggle">Voulez-vous sauvegarder les photos ?</label>
<input type="checkbox" id="save-toggle" class="toggle">
"""

css = '''
<style>

.chat-message {
    padding: 1.5rem; border-radius: 0.5rem; margin-bottom: 1rem; display: flex
}
.chat-message.user {
    background-color: #2b313e
}
.chat-message.bot {
    background-color: #475063
}
.chat-message .avatar {
  width: 20%;
}
.chat-message .avatar img {
  max-width: 78px;
  max-height: 78px;
  border-radius: 50%;
  object-fit: cover;
}
.chat-message .message {
  width: 80%;
  padding: 0 1.5rem;
  color: #fff;
}
.custom-header {
    color: white;
}

[data-testid="stFileUploadDropzone"] div div::before {color:red; content:"This text replaces Drag and drop file here"}
[data-testid="stFileUploadDropzone"] div div span{display:none;}
[data-testid="stFileUploadDropzone"] div div::after {color:red; font-size: .8em; content:"This text replaces Limit 200MB per file"}
[data-testid="stFileUploadDropzone"] div div small{display:none;}
</style>
'''


bot_template = '''
<div class="chat-message bot">
    <div class="avatar">
        <img src="https://i.ibb.co/cN0nmSj/Screenshot-2023-05-28-at-02-37-21.png" style="max-height: 78px; max-width: 78px; border-radius: 50%; object-fit: cover;">
    </div>
    <div class="message">{{MSG}}</div>
</div>
'''

user_template = '''
<div class="chat-message user">
    <div class="avatar">
        <img src="https://i.ibb.co/rdZC7LZ/Photo-logo-1.png">
    </div>    
    <div class="message">{{MSG}}</div>
</div>
'''
