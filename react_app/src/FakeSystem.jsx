import React, { useState, useEffect, useRef } from 'react';
import BatchWizardModal from './BatchWizardModal.jsx';
import TplConfigListModal from './TplConfigListModal.jsx';
import LetterEditorModal from './LetterEditorModal.jsx';
import QuickImportModal from './QuickImportModal.jsx';

/**
 * 模拟真实「函证系统」制函管理页面（参考用户截图 + docs/proto/集成方案3_定稿）。
 * - 顶部主 Tab 与筛选区、函证列表表格仅做展示；
 * - 真实交互：Excel制函（批量制函对话框）、模板配置列表弹窗 + 修改模板进入 OnlyOffice 编辑页。
 */

const TABS = [
  '控制表数据维护',
  '制函管理',
  '发函管理',
  '发函记录',
  '回函管理',
  '未回函管理',
  '回函过程检查',
  '已确认结果函证',
  '函证文件',
];

// 20 条演示数据，其中前 4 条与后端 ledger.json 对应，便于批量制函匹配测试
const TABLE_DATA = [
  { no: 'wlfz002001', cust: '中信证券股份有限公司', unit: '中信证券', addr: '广东省深圳市福田区中心...', contact: '王武', phone: '13800000000' },
  { no: 'wlfz002007', cust: '国泰君安资产管理有限公司', unit: '国泰君安', addr: '中国（上海）自由贸易试...', contact: '李明', phone: '13800000001' },
  { no: 'wlfz002013', cust: '华泰证券股份有限公司', unit: '华泰证券', addr: '江苏省南京市江东中路22...', contact: '张伟', phone: '13800000002' },
  { no: 'wlfz002021', cust: '广发基金管理有限公司', unit: '广发基金', addr: '广东省广州市黄埔区中新...', contact: '赵敏', phone: '13800000003' },
  { no: 'wlfz002002', cust: '顶点软件', unit: '国泰君安证券股份有限公司', addr: '中国（上海）自由贸易试...', contact: '李明', phone: '13800000011' },
  { no: 'wlfz002003', cust: '顶点软件', unit: '华泰证券股份有限公司', addr: '江苏省南京市江东中路22...', contact: '张伟', phone: '13800000012' },
  { no: 'wlfz002004', cust: '顶点软件', unit: '招商证券股份有限公司', addr: '广东省深圳市福田区福田...', contact: '王芳', phone: '13800000013' },
  { no: 'wlfz002005', cust: '顶点软件', unit: '海通证券股份有限公司', addr: '上海市广东路689号', contact: '刘洋', phone: '13800000014' },
  { no: 'wlfz002006', cust: '顶点软件', unit: '广发证券股份有限公司', addr: '广东省广州市黄埔区中新...', contact: '陈静', phone: '13800000015' },
  { no: 'wlfz002007', cust: '顶点软件', unit: '中国国际金融股份有限公司', addr: '北京市朝阳区建国门外大...', contact: '赵强', phone: '13800000016' },
  { no: 'wlfz002008', cust: '顶点软件', unit: '申万宏源证券有限公司', addr: '上海市徐汇区长乐路989...', contact: '孙杰', phone: '13800000017' },
  { no: 'wlfz002009', cust: '顶点软件', unit: '光大证券股份有限公司', addr: '上海市静安区新闸路150...', contact: '周婷', phone: '13800000018' },
  { no: 'wlfz002010', cust: '顶点软件', unit: '平安证券股份有限公司', addr: '广东省深圳市福田区益田...', contact: '吴磊', phone: '13800000019' },
  { no: 'wlfz002011', cust: '顶点软件', unit: '东方证券股份有限公司', addr: '上海市黄浦区中山南路31...', contact: '郑华', phone: '13800000020' },
  { no: 'wlfz002012', cust: '顶点软件', unit: '兴业证券股份有限公司', addr: '福建省福州市湖东路26...', contact: '钱坤', phone: '13800000021' },
  { no: 'wlfz002014', cust: '顶点软件', unit: '国信证券股份有限公司', addr: '广东省深圳市罗湖区红岭...', contact: '冯鑫', phone: '13800000022' },
  { no: 'wlfz002015', cust: '顶点软件', unit: '中国银河证券股份有限公司', addr: '北京市丰台区西营街8号...', contact: '许强', phone: '13800000023' },
  { no: 'wlfz002016', cust: '顶点软件', unit: '安信证券股份有限公司', addr: '广东省深圳市福田区福华...', contact: '何静', phone: '13800000024' },
  { no: 'wlfz002017', cust: '顶点软件', unit: '方正证券股份有限公司', addr: '湖南省长沙市天心区湘江...', contact: '林涛', phone: '13800000025' },
  { no: 'wlfz002018', cust: '顶点软件', unit: '长江证券股份有限公司', addr: '湖北省武汉市江汉区新华...', contact: '黄磊', phone: '13800000026' },
];

function Dropdown({ label, items }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  // 点击外部关闭
  useEffect(() => {
    if (!open) return;
    const onDoc = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, [open]);

  return (
    <div className="fake-dd-wrap" ref={ref}>
      <button className="fake-btn" onClick={() => setOpen((v) => !v)}>
        {label} <span className="fake-arrow">▾</span>
      </button>
      {open && (
        <div className="fake-dd">
          {items.map((it, idx) =>
            it.sep ? (
              <div key={idx} className="fake-dd-sep" />
            ) : (
              <div key={idx} className="fake-dd-item" onClick={() => { setOpen(false); it.onClick && it.onClick(); }}>
                <span className="fake-dd-icon">{it.icon}</span>
                {it.label}
              </div>
            )
          )}
        </div>
      )}
    </div>
  );
}

export default function FakeSystem() {
  const [activeTab, setActiveTab] = useState('制函管理');
  // 弹窗类型：null | {type:'batch'} | {type:'tpllist'} | {type:'tpledit', template}
  const [modal, setModal] = useState(null);

  return (
    <div className="fake-system">
      {/* 顶部主 Tab */}
      <div className="fake-header">
        <div className="fake-header-top">
          <span>往来函证0813</span>
          <span>演示环境 · 管理员</span>
        </div>
        <div className="fake-tabs">
          {TABS.map((t) => (
            <div
              key={t}
              className={`fake-tab ${activeTab === t ? 'active' : ''}`}
              onClick={() => setActiveTab(t)}
            >
              {t}
            </div>
          ))}
        </div>
      </div>

      {/* 主内容 */}
      {activeTab === '制函管理' ? (
        <div className="fake-page">
          <h1 className="fake-page-title">函证管理</h1>

          {/* 筛选区 */}
          <div className="fake-filter">
            <div className="fake-filter-row">
              <label>关键字：</label>
              <input type="text" placeholder="被询证单位、函证编号、被询证单位地址" />

              <label>是否制函：</label>
              <div className="fake-radio"><input type="radio" name="made" defaultChecked /> 全部</div>
              <div className="fake-radio"><input type="radio" name="made" /> 是</div>
              <div className="fake-radio"><input type="radio" name="made" /> 否</div>

              <label>是否发函：</label>
              <div className="fake-radio"><input type="radio" name="sent" defaultChecked /> 全部</div>
              <div className="fake-radio"><input type="radio" name="sent" /> 是</div>
              <div className="fake-radio"><input type="radio" name="sent" /> 否</div>

              <label>是否废弃：</label>
              <div className="fake-radio"><input type="radio" name="discard" defaultChecked /> 全部</div>
              <div className="fake-radio"><input type="radio" name="discard" /> 是</div>
              <div className="fake-radio"><input type="radio" name="discard" /> 否</div>
            </div>
            <div className="fake-filter-row">
              <label>制函日期：</label>
              <input type="text" placeholder="开始日期" className="date" /> →
              <input type="text" placeholder="结束日期" className="date" />

              <label>发函日期：</label>
              <input type="text" placeholder="开始日期" className="date" /> →
              <input type="text" placeholder="结束日期" className="date" />

              <button className="fake-btn primary">查询</button>
              <button className="fake-btn">重置</button>
            </div>
          </div>

          {/* 操作按钮区（定稿布局：Excel制函往前放）*/}
          <div className="fake-btn-row">
            <button className="fake-btn">批量上传原始询证函</button>

            <Dropdown
              label="Excel制函"
              items={[
                { icon: '📥', label: '批量制函（Excel）', onClick: () => setModal({ type: 'batch' }) },
                { sep: true },
                { icon: '🧩', label: '模板配置', onClick: () => setModal({ type: 'tpllist' }) },
                { icon: '⚡', label: '快速导入Excel制函', onClick: () => setModal({ type: 'quick' }) },
              ]}
            />

            <Dropdown
              label="标准科目函证制函"
              items={[
                { icon: '📄', label: 'PDF格式制函', onClick: () => alert('PDF格式制函为演示占位功能') },
                { icon: '🛠️', label: '修改制函模板', onClick: () => setModal({ type: 'tpllist' }) },
              ]}
            />

            <button className="fake-btn">批量合并下载</button>
            <button className="fake-btn">数据修改日志</button>
          </div>

          {/* 函证列表 */}
          <table className="fake-table">
            <thead>
              <tr>
                <th>序号</th>
                <th>函证编号</th>
                <th>客户名称</th>
                <th>被询证单位</th>
                <th>被询证单位地址</th>
                <th>被询证单位联系人</th>
                <th>被询证单位电话</th>
                <th>操作</th>
              </tr>
            </thead>
            <tbody>
              {TABLE_DATA.map((row, idx) => (
                <tr key={idx}>
                  <td>{idx + 1}</td>
                  <td>{row.no}</td>
                  <td>{row.cust}</td>
                  <td>{row.unit}</td>
                  <td>{row.addr}</td>
                  <td>{row.contact}</td>
                  <td>{row.phone}</td>
                  <td>
                    <span className="fake-op">上传原始询证函</span>
                    <span className="fake-op">上传附表</span>
                    <span className="fake-op">下载标记询证函</span>
                    <span className="fake-op">修改函证数据</span>
                    <span className="fake-op del">废弃</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* 分页 */}
          <div className="fake-pagination">
            <span>共 20 条</span>
            <button className="fake-page-btn">&lt;</button>
            <button className="fake-page-btn active">1</button>
            <button className="fake-page-btn">2</button>
            <button className="fake-page-btn">&gt;</button>
            <span>10条/页</span>
            <span className="fake-jump">跳至 <input type="text" /> 页</span>
          </div>
        </div>
      ) : (
        <div className="fake-placeholder">
          <div className="fake-placeholder-icon">🚧</div>
          <div className="fake-placeholder-title">{activeTab}</div>
          <div className="fake-placeholder-desc">
            该模块为演示占位，暂无实际功能。<br />
            真实系统对接后此处将展示相应的业务界面。
          </div>
        </div>
      )}

      {/* 真实弹窗 */}
      {modal && modal.type === 'batch' && (
        <BatchWizardModal onClose={() => setModal(null)} />
      )}
      {modal && modal.type === 'tpllist' && (
        <TplConfigListModal
          onClose={() => setModal(null)}
          onEditTemplate={(t) => setModal({ type: 'tpledit', template: t })}
        />
      )}
      {modal && modal.type === 'tpledit' && (
        <LetterEditorModal template={modal.template} onClose={() => setModal({ type: 'tpllist' })} />
      )}
      {modal && modal.type === 'quick' && (
        <QuickImportModal onClose={() => setModal(null)} />
      )}
    </div>
  );
}
