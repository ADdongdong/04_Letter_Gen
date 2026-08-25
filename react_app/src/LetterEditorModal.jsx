import React, { useState, useRef, useEffect } from 'react';
import OnlyOfficeEditor from './OnlyOfficeEditor.jsx';
import RightPanel from './RightPanel.jsx';

/**
 * 模板编辑页（OnlyOffice）：从模板配置列表弹窗点「修改模板」进入。
 * - template 为 null → 新建模板，左侧为空，等待用户上传 Word；
 * - template 有 id → 打开该模板 URL（/api/oo/template/<id>）进行编辑/保存占位符绑定。
 * 相对 URL 由 OnlyOfficeEditor 统一转成容器可访问的 host.docker.internal:5002 绝对地址。
 */
export default function LetterEditorModal({ template, onClose }) {
  const isEdit = !!(template && template.id);
  const initUrl = isEdit ? '/api/oo/template/' + template.id : '/api/oo/template/blank';

  const [curDocKey, setCurDocKey] = useState(`cfg-${isEdit ? template.id : 'new'}-` + new Date().toISOString().slice(0, 10));
  const [curDocUrl, setCurDocUrl] = useState(initUrl);
  const [sourceId, setSourceId] = useState(isEdit ? template.id : 'blank'); // 当前基于哪个 docx（default/blank/模板id）
  const [editId, setEditId] = useState(isEdit ? template.id : null);        // 修改模式：被更新的模板 id
  const [templates, setTemplates] = useState([]);
  const [status, setStatus] = useState('就绪');
  const [uploadTip, setUploadTip] = useState(''); // 上传 Word 模板成功提示（优先于普通状态展示）
  const [editorReady, setEditorReady] = useState(false);
  const editorRef = useRef(null);

  const updateStatus = (msg) => setStatus(msg);
  const loadTemplates = async () => {
    try {
      const res = await fetch('/api/templates');
      const data = await res.json();
      setTemplates(data.templates || []);
    } catch (e) { /* ignore */ }
  };
  useEffect(() => { loadTemplates(); }, []);

  const onReady = () => {
    updateStatus('编辑器就绪');
    setEditorReady(true);
  };
  const insertText = (text) => {
    if (!editorRef.current) return false;
    try { return editorRef.current.insertText(text); }
    catch (e) { return false; }
  };

  // 上传/替换 Word 模板 → 后端保存 → 左侧 OnlyOffice 打开
  const onUploadTemplate = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append('file', file);
    try {
      let res, data;
      if (isEdit) {
        // 修改模式：替换当前模板 Word 文件，不新建清单条目
        fd.append('replace_id', editId);
        res = await fetch('/api/upload_template_replace', { method: 'POST', body: fd });
        data = await res.json();
        if (data.error) { alert('替换失败：' + data.error); return; }
        setCurDocUrl(data.url);
        setCurDocKey('rep-' + data.id + '-' + new Date().toISOString().slice(11, 19).replace(/:/g, ''));
        const tip = '✓ 已替换 Word 模板：' + data.name;
        setUploadTip(tip);
        updateStatus(tip);
      } else {
        // 新建模式：上传后生成临时模板条目，保存时原地更新
        res = await fetch('/api/upload_template_oo', { method: 'POST', body: fd });
        data = await res.json();
        if (data.error) {
          const msg = res.status === 409 ? '模板名称重复：' + data.error : '上传失败：' + data.error;
          alert(msg);
          return;
        }
        setEditId(data.id);
        setSourceId(data.id);
        setCurDocUrl(data.url);
        setCurDocKey('up-' + data.id + '-' + new Date().toISOString().slice(11, 19).replace(/:/g, ''));
        await loadTemplates();
        const tip = '✓ 已加载上传的 Word 模板：' + data.name;
        setUploadTip(tip);
        updateStatus(tip);
      }
    } catch (err) {
      alert(isEdit ? '替换失败：' + err.message : '上传失败：' + err.message);
    } finally {
      e.target.value = '';
    }
  };

  return (
    <div className="editor-page">
      <div className="editor-page-header">
        <span className="editor-page-title">{template ? '修改制函模板 · ' + template.name : '新建制函模板'}</span>
        <button className="editor-page-back" onClick={onClose}>← 返回</button>
      </div>
      <div className="editor-page-body">
        <div className="editor-workbench">
          <div className="onlyoffice-box">
            <OnlyOfficeEditor
              ref={editorRef}
              key={curDocKey}
              docKey={curDocKey}
              docUrl={curDocUrl}
              onReady={onReady}
            />
          </div>
          <RightPanel
            ready={editorReady}
            insertText={insertText}
            updateStatus={updateStatus}
            templates={templates}
            onLoadTemplates={loadTemplates}
            onOpenTemplate={(url) => setCurDocUrl(url)}
            defaultTab="config"
            showTabs={false}
            hideSourceSelect={true}
            onUploadTemplate={onUploadTemplate}
            sourceTemplateId={sourceId}
            mode={template ? 'edit' : 'new'}
            templateName={template ? template.name : ''}
            templateId={editId}
            />
        </div>
      </div>
      <div className="editor-page-status">状态：{uploadTip || status}</div>
    </div>
  );
}
