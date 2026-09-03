import React, { useState, useRef, useEffect } from 'react';

/**
 * 自绘下拉选择组件（替代原生 select）：
 * - 触发器与输入框同款视觉（白底细边 8px 圆角，focus 蓝环，见 styles.css .sfield-trigger）
 * - 弹出面板贴合主题（白底圆角阴影、选项 hover 浅蓝、选中主色，见 .sfield-menu/.sfield-opt）
 * - 根除原生 select 弹层无法定制样式 + 收起闪烁的问题
 *
 * Props:
 *   value     当前选中值
 *   onChange  (value) => void
 *   options   [{ value, label }]
 *   disabled  禁用态（可选）
 *   style     外层容器附加样式（可选）
 */
export default function SelectField({ value, onChange, options, disabled, style }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);

  // 点击组件外部时收起面板（open 时才注册监听）
  useEffect(() => {
    if (!open) return;
    const onDocMouseDown = (e) => {
      if (rootRef.current && !rootRef.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [open]);

  const current = options.find(o => o.value === value);
  const pick = (v) => { onChange(v); setOpen(false); };

  return (
    <div ref={rootRef} style={{ position: 'relative', ...style }}>
      <button
        type="button"
        className="sfield-trigger"
        disabled={disabled}
        onClick={() => setOpen(o => !o)}
      >
        <span>{current ? current.label : '请选择'}</span>
        <span className={open ? 'sfield-arrow open' : 'sfield-arrow'}>▾</span>
      </button>
      {open && (
        <div className="sfield-menu">
          {options.map(o => (
            <div
              key={String(o.value)}
              className={o.value === value ? 'sfield-opt active' : 'sfield-opt'}
              onClick={() => pick(o.value)}
            >
              {o.label}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
