import React from "react";

/**
 * UI primitives to enforce layout invariants across the dialog.
 * Goal: make list containment + row truncation impossible to "drift" across tabs.
 */

export function ListBox({ height, style, children, ...rest }) {
  const h = typeof height === "number" ? `${height}px` : height;
  return (
    <div
      {...rest}
      style={{
        height: h,
        maxHeight: h,
        minHeight: h,
        maxWidth: "100%",
        minWidth: 0,
        boxSizing: "border-box",
        overscrollBehavior: "contain",
        overflowY: "auto",
        overflowX: "hidden",
        border: "1px solid rgba(0,0,0,0.1)",
        borderRadius: 6,
        ...style,
      }}
    >
      {children}
    </div>
  );
}

export function RowLabel({ style, children, title, ...rest }) {
  const text = typeof children === "string" ? children : undefined;
  return (
    <div
      {...rest}
      title={title ?? text}
      style={{
        flex: "1 1 auto",
        minWidth: 0,
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
        ...style,
      }}
    >
      {children}
    </div>
  );
}
