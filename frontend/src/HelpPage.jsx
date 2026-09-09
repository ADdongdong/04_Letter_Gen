import React from 'react';

/**
 * 使用说明页（#/help）：说明制函功能的应对场景与"配置 → 标注 → 保存 → 制函"全流程。
 * 面向首次使用者：先讲清解决什么问题，再按步骤走通，最后列常见规则。
 */
export default function HelpPage() {
  return (
    <div className="page" style={{ maxWidth: '860px', margin: '0 auto' }}>
      <h2 style={{ margin: '8px 0 16px' }}>使用说明：从模板配置到批量制函</h2>

      <div className="card">
        <h3>🎯 这个功能解决什么问题？</h3>
        <p style={{ lineHeight: 1.8 }}>
          审计人员出具往来函证时，需要把 Excel 里的数据（科目余额、非标事项、票据等）
          以表格形式填入 Word 函证。手工逐表复制粘贴<b>效率低、易错位</b>。
        </p>
        <p style={{ lineHeight: 1.8 }}>
          本功能把这件事变成两步：<b>① 用 Word 模板定版式、标注表格插入位置</b>；
          <b>② 用 Excel 提供数据，一键批量生成全部函证（ZIP 打包下载）</b>。
          一份 Excel 可以同时包含多封函证的数据，系统按
          「被审计单位 + 被询证单位」自动拆分成 N 封独立文档。
        </p>
        <div className="hint">
          适用范围：批量制函仅针对<b>往来函证</b>（被审计单位 + 被询证单位模式），不针对银行函证。
        </div>
      </div>

      <div className="card" style={{ marginTop: '12px' }}>
        <h3>📝 第一步：配置制函模板（一次配置，反复使用）</h3>
        <ol style={{ lineHeight: 2 }}>
          <li>进入<b>模板列表页</b>，点右上角【新增模板】；</li>
          <li>
            在右侧上传 <b>Word 模板</b>（函证版式：抬头、正文、落款等固定内容——
            可先 <a href="/api/templates/sample/word" download>下载示例 Word 模板</a> 修改使用）；
          </li>
          <li>
            上传 <b>Excel</b>（只需要 Sheet 名称与表头结构，<b>可以没有数据</b>——
            它告诉系统"这份模板会用到哪些 Sheet"）；
          </li>
          <li>
            在左侧编辑器中把<b>光标点到希望插入表格的位置</b>，
            再点击右侧对应的 Sheet 名称 → 系统即在光标处插入一行蓝色占位标注
            （形如 <code>【Sheet「科目余额分录」表格将在此处展示】</code>）；
            每个要出表格的位置都标注一次，同一 Sheet 可在多处插入；
          </li>
          <li>填写模板名称，点【保存模板配置】——系统自动把编辑器当前内容（含标注）回写保存。</li>
        </ol>
        <div className="hint">修改模板：列表中点【修改】加载已保存的 Word，改完保存即覆盖。</div>
      </div>

      <div className="card" style={{ marginTop: '12px' }}>
        <h3>🚀 第二步：Excel 批量制函（日常使用）</h3>
        <ol style={{ lineHeight: 2 }}>
          <li>进入<b>「Excel 制函」页</b>，下拉选择一个制函模板；</li>
          <li>
            上传<b>批量 Excel</b>
            （可先 <a href="/api/templates/sample/excel" download>下载数据导入模板</a>，
            按示例填入实际数据）：
            <ul>
              <li>各 Sheet 表头<b>前两列必须为「被审计单位」「被询证单位」</b>——这两列的组合唯一确定一封函证；</li>
              <li>制函生成的表格从第 3 列开始（前两列只用于分组与文件命名，不出现在表格里）；</li>
              <li>同一封函证的数据可分散在多个 Sheet，系统自动归集。</li>
            </ul>
          </li>
          <li>
            上传后系统<b>即时校验与匹配</b>：格式不符、分组列为空会精确提示到 Sheet/行/列；
            并把单位组合与<b>函证系统登记数据</b>比对——匹配成功带出函证编号，
            <span className="warn-red">未匹配的组合（红色行）整封跳过</span>；
          </li>
          <li>点【制函】→ 每封函证独立渲染 → ZIP 自动下载。</li>
        </ol>
      </div>

      <div className="card" style={{ marginTop: '12px' }}>
        <h3>📌 关键规则速查</h3>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', lineHeight: 1.8 }}>
          <tbody>
            <tr style={{ background: 'var(--primary-bg)' }}>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px', width: '220px' }}><b>占位段</b></td>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}>
                制函时被替换为真实表格；某 Sheet 无数据或 Excel 中不存在时，该占位段<b>自动删除</b>（不残留提示文字）
              </td>
            </tr>
            <tr>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}><b>函证编号</b></td>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}>
                仅系统已登记的函证会生成；全部未登记时无法制函，需先在函证系统完成登记
              </td>
            </tr>
            <tr style={{ background: 'var(--primary-bg)' }}>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}><b>表格样式</b></td>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}>
                默认开箱即用（10pt 居中网格）；可在配置页「表格样式（高级）」中调整字号/字体/对齐/行高/列宽
              </td>
            </tr>
            <tr>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}><b>模板即定义</b></td>
              <td style={{ border: '1px solid #ddd', padding: '6px 10px' }}>
                制函时的插入关系由模板中的占位段自动提取，无需在制函页重复设置
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div className="card" style={{ marginTop: '12px', marginBottom: '20px' }}>
        <h3>✅ 建议的上手路径</h3>
        <p style={{ lineHeight: 1.8 }}>
          下载示例 Word 模板 → 新增模板上传它 + 上传一份 Sheet 结构 Excel → 随便标注一个位置 →
          保存 → 到制函页下载数据导入模板 → 原样上传 → 点制函 → 打开 ZIP 看效果。
          全程约 5 分钟，即可理解全部流程。
        </p>
      </div>
    </div>
  );
}
