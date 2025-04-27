import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import seaborn as sns
import math

# — Global state —
sources = {}      # file path → { type:'excel'|'csv', xls:ExcelFile, sheets:[…] }
y_axes = []       # list of Y‑axis configurations
fig = None
canvas_plot = None
saved_plots = [] # ← holds dicts of {'x':…, 'ys':…,'x_label':…,'title':…}

# — Helper functions —
def load_files():
    paths = filedialog.askopenfilenames(
        filetypes=[
            ("Excel files", ("*.xlsx","*.xls")),
            ("CSV files",   ("*.csv",)),
            ("All files",   ("*.*",))
        ]
    )
    if not paths:
        return
    sources.clear()
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        try:
            if ext in ('.xls','.xlsx'):
                xls = pd.ExcelFile(p)
                sources[p] = {'type':'excel', 'xls':xls, 'sheets':xls.sheet_names}
            else:
                sources[p] = {'type':'csv',   'sheets':['(csv)']}
        except Exception as e:
            messagebox.showerror("Load error", f"Cannot load {p}:\n{e}")
    all_files = list(sources.keys())
    file_x_cb['values'] = all_files
    for cfg in y_axes:
        for cb in cfg['file_cbs']:
            cb['values'] = all_files

def get_df(path, sheet):
    info = sources.get(path)
    if not info:
        raise ValueError("No file selected")
    if info['type']=='excel':
        return info['xls'].parse(sheet)
    else:
        return pd.read_csv(path)

def make_file_handler(file_cb, sheet_cb, col_cbs):
    def on_file(_=None):
        path = file_cb.get()
        info = sources.get(path)
        if not info:
            return
        sheet_cb['values'] = info['sheets']
        sheet_cb.set(info['sheets'][0])
        on_sheet()
    def on_sheet(_=None):
        try:
            df = get_df(file_cb.get(), sheet_cb.get())
        except Exception as e:
            messagebox.showerror("Read error", str(e))
            return
        cols = list(df.columns)
        for cb in col_cbs:
            cb['values'] = cols
            cb.set('')
    return on_file, on_sheet

def compute_series(cfg):
    """Compute a pandas Series for one Y‑axis config."""
    if cfg['type_var'].get()=='direct':
        f_cb, s_cb, c_cb = cfg['direct']
        df = get_df(f_cb.get(), s_cb.get())
        return df[c_cb.get()]
    # derived
    f1, s1, c1, op_cb, f2, s2, c2 = cfg['derived']
    df1 = get_df(f1.get(), s1.get())
    df2 = get_df(f2.get(), s2.get())
    sA, sB = df1[c1.get()], df2[c2.get()]
    op = op_cb.get()
    if op=='+': return sA + sB
    if op=='-': return sA - sB
    if op=='*': return sA * sB
    if op=='/': return sA / sB
    raise ValueError("Select an operation")

# — Y‑Axis management —
def add_y_axis():
    idx = len(y_axes)
    frame = ttk.LabelFrame(y_container, text=f"Y‑Axis #{idx+1}", padding=8)
    frame.grid(row=idx, column=0, pady=5, sticky="we")

    # Direct vs Derived
    type_var = tk.StringVar(value='direct')
    ttk.Radiobutton(frame, text="Direct",  variable=type_var, value='direct',
                    command=lambda i=idx: update_one(i)).grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(frame, text="Derived", variable=type_var, value='derived',
                    command=lambda i=idx: update_one(i)).grid(row=0, column=1, sticky="e")

    # Direct config
    direct_frame = ttk.Frame(frame)
    direct_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=2)
    file_cb  = ttk.Combobox(direct_frame, state='readonly')
    sheet_cb = ttk.Combobox(direct_frame, state='readonly')
    col_cb   = ttk.Combobox(direct_frame, state='readonly')
    for r, lbl in enumerate(("File:","Sheet:","Col:")):
        ttk.Label(direct_frame, text=lbl).grid(row=r, column=0, sticky="w")
    file_cb.grid(row=0, column=1, sticky="we")
    sheet_cb.grid(row=1, column=1, sticky="we")
    col_cb.grid(row=2, column=1, sticky="we")

    # Derived config
    derived_frame = ttk.Frame(frame)
    derived_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=2)
    f1 = ttk.Combobox(derived_frame, state='readonly'); s1 = ttk.Combobox(derived_frame, state='readonly')
    c1 = ttk.Combobox(derived_frame, state='readonly'); op = ttk.Combobox(derived_frame, values=['+','-','*','/'], state='readonly')
    f2 = ttk.Combobox(derived_frame, state='readonly'); s2 = ttk.Combobox(derived_frame, state='readonly')
    c2 = ttk.Combobox(derived_frame, state='readonly')
    for r, lbl in enumerate(("File1:","Sheet1:","Col1:","Op:","File2:","Sheet2:","Col2:")):
        ttk.Label(derived_frame, text=lbl).grid(row=r, column=0, sticky="w")
    f1.grid(row=0, column=1, sticky="we");  s1.grid(row=1, column=1, sticky="we")
    c1.grid(row=2, column=1, sticky="we");  op.grid(row=3, column=1, sticky="we")
    f2.grid(row=4, column=1, sticky="we");  s2.grid(row=5, column=1, sticky="we")
    c2.grid(row=6, column=1, sticky="we")

    # Label & Remove
    lbl_entry = ttk.Entry(frame)
    ttk.Label(frame, text="Label:").grid(row=3, column=0, sticky="w")
    lbl_entry.grid(row=3, column=1, sticky="we")

    # ── AUTO‑SET LABEL TO THE SELECTED COLUMN ──
    col_cb.bind("<<ComboboxSelected>>", lambda e, lb=lbl_entry, cb=col_cb: (
        lb.delete(0, tk.END),
        lb.insert(0, cb.get())
    ))

    remove_btn = ttk.Button(frame, text="Remove", command=lambda i=idx: remove_y_axis(i))
    remove_btn.grid(row=4, column=0, columnspan=2, pady=4)

    # Bind file/sheet handlers
    for fcb, scb, col_list in [(file_cb, sheet_cb, [col_cb]),
                                (f1, s1, [c1]), (f2, s2, [c2])]:
        on_f, on_s = make_file_handler(fcb, scb, col_list)
        fcb.bind("<<ComboboxSelected>>", on_f)
        scb.bind("<<ComboboxSelected>>", on_s)

    cfg = {
        'frame': frame,
        'type_var': type_var,
        'direct_frame': direct_frame,
        'derived_frame': derived_frame,
        'direct': (file_cb, sheet_cb, col_cb),
        'derived': (f1, s1, c1, op, f2, s2, c2),
        'file_cbs': [file_cb, f1, f2],
        'label': lbl_entry
    }
    y_axes.append(cfg)
    all_files = list(sources.keys())
    for cb in cfg['file_cbs']:
        cb['values'] = all_files

    update_one(idx)

def update_one(idx):
    cfg = y_axes[idx]
    if cfg['type_var'].get()=='direct':
        cfg['direct_frame'].grid()
        cfg['derived_frame'].grid_remove()
    else:
        cfg['direct_frame'].grid_remove()
        cfg['derived_frame'].grid()

def remove_y_axis(idx):
    cfg = y_axes.pop(idx)
    cfg['frame'].destroy()
    # re-index remaining
    for i, c in enumerate(y_axes):
        c['frame'].grid_configure(row=i)
        c['frame'].configure(text=f"Y‑Axis #{i+1}")
        for w in c['frame'].winfo_children():
            if isinstance(w, ttk.Button) and w['text']=="Remove":
                w.configure(command=lambda i=i: remove_y_axis(i))

    # ── Seaborn theme for grid & palette ──
    sns.set_theme(style="darkgrid", palette="deep", font_scale=1.0)
    sns.set_context("talk")

    # ── Load X data ──
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Compute Y series & labels ──
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    if not ys:
        return messagebox.showwarning("No data", "Add at least one Y‑axis.")

    # ── Create or clear figure ──
    if not fig:
        fig = plt.Figure(figsize=(8, 6), tight_layout=True)
    else:
        fig.clear()

    host = fig.add_subplot(111)

    # ── Pick two colors for left & right axes ──
    palette = sns.color_palette("deep", n_colors=2)
    left_color, right_color = palette[0], palette[1]

    # ── Plot first series on left Y‑axis ──
    line1, = host.plot(
        x, ys[0],
        label=labels[0],
        color=left_color,
        marker='o', markersize=6,
        linewidth=2, alpha=0.9
    )
    host.set_ylabel(labels[0], color=left_color, labelpad=15)
    host.tick_params(axis='y', colors=left_color)

    # ── Plot second series (if any) on right Y‑axis ──
    ax2 = None
    if len(ys) > 1:
        ax2 = host.twinx()
        line2, = ax2.plot(
            x, ys[1],
            label=labels[1],
            color=right_color,
            marker='s', markersize=6,
            linewidth=2, alpha=0.9
        )
        ax2.set_ylabel(labels[1], color=right_color, labelpad=15)
        ax2.tick_params(axis='y', colors=right_color)

    # ── Labels & title ──
    host.set_xlabel(x_label.get(), fontsize=12, labelpad=15)
    fig.suptitle(title_entry.get(), fontsize=14, y=0.98)

    # ── Adjust right margin for legend ──
    fig.subplots_adjust(right=0.75)

    # ── Legend outside to the right ──
    handles = [line1]
    legend_labels = [labels[0]]
    if ax2:
        handles.append(line2)
        legend_labels.append(labels[1])
    host.legend(
        handles, legend_labels,
        bbox_to_anchor=(1.02, 1),
        loc='upper left',
        borderaxespad=0.
    )

    # ── Draw in Tkinter ──
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)



    import seaborn as sns
    # ── Seaborn theming ──
    sns.set_theme(style="darkgrid", palette="deep", font_scale=1.1)
    sns.set_context("talk")

    # ── Fetch X data ──
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Fetch Y series & labels ──
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    # ── Create or clear the figure ──
    if not fig:
        fig = plt.Figure(figsize=(8, 6), tight_layout=True)
    else:
        fig.clear()

    host = fig.add_subplot(111)

    # ── Plot all series on the same Y‑axis ──
    palette = sns.color_palette("deep", n_colors=len(ys))
    markers = ['o','s','D','^','v','P','*']
    for y, lbl, color, m in zip(ys, labels, palette, markers):
        sns.lineplot(
            x=x, y=y, ax=host,
            label=lbl,
            color=color,
            marker=m, markersize=6,
            linewidth=2, alpha=0.9
        )

    # ── Axis labels & title ──
    host.set_xlabel(x_label.get(), fontsize=12, labelpad=15)
    host.set_ylabel("Value", fontsize=12, labelpad=15)
    fig.suptitle(title_entry.get(), fontsize=14, y=0.98)

    # ── Legend outside to the right ──
    fig.subplots_adjust(right=0.75)
    host.legend(
        bbox_to_anchor=(1.02, 1),
        loc='upper left',
        borderaxespad=0.
    )

    # ── Render into Tkinter ──
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)

    import seaborn as sns
    # ── Seaborn theming ──
    sns.set_theme(style="darkgrid", palette="deep", font_scale=1.1)
    sns.set_context("talk")

    # ── Fetch X data ──
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Fetch Y series & labels ──
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    # ── Create or clear the figure ──
    if not fig:
        fig = plt.Figure(figsize=(8, 6), tight_layout=True)
    else:
        fig.clear()

    host = fig.add_subplot(111)
    axes = [host]

    # ── Add extra Y‑axes ──
    for i in range(1, len(ys)):
        ax = host.twinx()
        ax.spines["right"].set_position(("axes", 1 + 0.1*(i-1)))
        ax.set_frame_on(True)
        ax.patch.set_visible(False)
        axes.append(ax)

    # ── Generate a color palette ──
    palette = sns.color_palette("deep", n_colors=len(axes))
    markers = ['o','s','D','^','v','P','*']

    # ── Plot each series, color + style its axis ──
    for idx, (ax, y, lbl, m, color) in enumerate(zip(axes, ys, labels, markers, palette)):
        sns.lineplot(
            x=x, y=y, ax=ax,
            color=color, marker=m, markersize=6,
            linewidth=2, alpha=0.9
        )

        # color spine & ticks & label
        side = "left" if idx == 0 else "right"
        ax.spines[side].set_color(color)
        ax.tick_params(axis='y', colors=color)
        ax.yaxis.label.set_color(color)

        # move label outside plot
        ax.set_ylabel(lbl, fontsize=11, labelpad=15)

        # endpoint annotation in matching color
        ax.annotate(
            f"{y.iloc[-1]:.2f}",
            xy=(x.iloc[-1], y.iloc[-1]),
            xytext=(5, 0), textcoords="offset points",
            va="center", fontsize=9, color=color
        )

    # ── X‑axis label ──
    host.set_xlabel(x_label.get(), fontsize=12, labelpad=15)
    # ── Title up top ──
    fig.suptitle(title_entry.get(), fontsize=14, y=0.98)

    # ── Make room on the right for legend ──
    fig.subplots_adjust(right=0.75)

    # ── Combine & place legend outside ──
    all_h, all_l = [], []
    for ax in axes:
        h, l = ax.get_legend_handles_labels()
        all_h += h; all_l += l
    host.legend(
        all_h, all_l,
        bbox_to_anchor=(1.02, 1),
        loc='upper left',
        borderaxespad=0.
    )

    # ── Render into Tkinter ──
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)


    # — Bring in Seaborn for theming (you still need seaborn installed) —
    import seaborn as sns
    sns.set_theme(style="darkgrid", palette="deep", font_scale=1.1)
    sns.set_context("talk")

    # — Grab X data —
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # — Compute all Y series and their labels —
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    # — Create or clear the figure —
    if not fig:
        fig = plt.Figure(figsize=(8, 6), tight_layout=True)
    else:
        fig.clear()

    host = fig.add_subplot(111)
    axes = [host]

    # — Add extra Y‑axes if needed —
    for i in range(1, len(ys)):
        ax = host.twinx()
        ax.spines["right"].set_position(("axes", 1 + 0.1*(i-1)))
        ax.set_frame_on(True)
        ax.patch.set_visible(False)
        axes.append(ax)

    # — Plot each series with markers & endpoint labels —
    markers = ['o','s','D','^','v','P','*']
    for ax, y, lbl, m in zip(axes, ys, labels, markers):
        sns.lineplot(
            x=x, y=y, ax=ax,
            label=lbl,
            marker=m, markersize=6,
            linewidth=2, alpha=0.85
        )
        ax.set_ylabel(lbl, fontsize=11)
        # annotate last point
        ax.annotate(
            f"{y.iloc[-1]:.2f}",
            xy=(x.iloc[-1], y.iloc[-1]),
            xytext=(5, 0),
            textcoords="offset points",
            va="center",
            fontsize=9
        )

    # — Final touches —
    host.set_xlabel(x_label.get(), fontsize=12)
    fig.suptitle(title_entry.get(), fontsize=14, y=0.98)

    # — Combine all legends into one —
    all_h, all_l = [], []
    for ax in axes:
        h, l = ax.get_legend_handles_labels()
        all_h += h; all_l += l
    host.legend(all_h, all_l, loc='best', frameon=True)

    # — Render into your Tk canvas —
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)

def save_current_plot_config():
    # ── Load X series ──
    try:
        df_x      = get_df(file_x_cb.get(), sheet_x_cb.get())
        x_series  = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Load all Y series & labels ──
    ys_list, labels = [], []
    for cfg in y_axes:
        try:
            ys_list.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    # ── DUPLICATE CHECK ──
    for existing in saved_plots:
        # same X?
        if not existing['x'].equals(x_series):
            continue
        # same number of Y series?
        if len(existing['ys_series']) != len(ys_list):
            continue
        # each Y identical?
        if any(not e_y.equals(y) for e_y, y in zip(existing['ys_series'], ys_list)):
            continue
        # same labels, x_label and title?
        if (existing['labels'] == labels and
            existing['x_label'] == x_label.get() and
            existing['title']   == title_entry.get()):
            return messagebox.showinfo("Duplicate", "That exact plot is already in your gallery.")

    # ── If we get here, it’s new — so save it ──
    saved_plots.append({
        'x':         x_series,
        'ys_series': ys_list,
        'labels':    labels,
        'x_label':   x_label.get(),
        'title':     title_entry.get()
    })

    messagebox.showinfo("Saved", f"Plot #{len(saved_plots)} added to gallery.")

def plot():
    global fig, canvas_plot      # must be first!


    # ── Theme & context ──
    sns.set_theme(style="whitegrid", palette="deep", font_scale=1.0)
    sns.set_context("talk")

    # ── Load X series ──
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Load all Y series & labels ──
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))
    n = len(ys)
    if n == 0:
        return messagebox.showwarning("No data", "Add at least one Y‑axis.")

    # ── Create a figure with n stacked subplots ──
    height = max(3, 2.5 * n)
    if not fig:
        fig = plt.Figure(figsize=(8, height), tight_layout=True)
    else:
        fig.clear()
    axs = fig.subplots(nrows=n, ncols=1, sharex=True)

    if n == 1:
        axs = [axs]  # make it iterable

    # ── Plot each series in its own subplot ──
    palette = sns.color_palette(n_colors=n)
    markers = ['o','s','D','^','v','P','*','X','d']
    for i, (ax, y, lbl) in enumerate(zip(axs, ys, labels)):
        color = palette[i]
        m = markers[i % len(markers)]
        ax.plot(
            x, y,
            label=lbl,
            color=color,
            marker=m, markersize=5,
            linewidth=2, alpha=0.85
        )
        ax.set_ylabel(lbl, color=color, fontsize=11)
        ax.tick_params(axis='y', colors=color, labelsize=9)
        ax.grid(alpha=0.3)

        # annotate last point
        ax.annotate(
            f"{y.iloc[-1]:.2f}",
            xy=(x.iloc[-1], y.iloc[-1]),
            xytext=(5, 5),
            textcoords="offset points",
            color=color,
            fontsize=9
        )

    # ── X label on bottom subplot ──
    axs[-1].set_xlabel(x_label.get(), fontsize=11)

    # ── Title and legend (one shared legend up top) ──
    fig.suptitle(title_entry.get(), fontsize=14, fontweight='bold', y=0.98)
    fig.legend(
        labels, loc='upper center',
        bbox_to_anchor=(0.5, 1.03),
        ncol=n if n<=4 else 4,
        fontsize='small',
        frameon=False
    )

    # ── Render to Tkinter ──
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)


    # ── Enriched Seaborn theme ──
    sns.set_theme(style="whitegrid", palette="muted", font_scale=1.1)
    sns.set_context("talk")

    # ── Fetch X data ──
    try:
        df_x = get_df(file_x_cb.get(), sheet_x_cb.get())
        x = df_x[col_x_cb.get()]
    except Exception as e:
        return messagebox.showerror("X error", str(e))

    # ── Compute all Y series & labels ──
    ys, labels = [], []
    for cfg in y_axes:
        try:
            ys.append(compute_series(cfg))
            labels.append(cfg['label'].get())
        except Exception as e:
            return messagebox.showerror("Y error", str(e))

    if not ys:
        return messagebox.showwarning("No data", "Add at least one Y‑axis.")

    # ── Create or clear the figure ──
    if not fig:
        fig = plt.Figure(figsize=(8, 6), tight_layout=True)
    else:
        fig.clear()
    host = fig.add_subplot(111)

    # ── Create additional Y‑axes ──
    axes = [host]
    for i in range(1, len(ys)):
        ax = host.twinx()
        ax.spines["right"].set_position(("axes", 1 + 0.1*(i-1)))
        ax.set_frame_on(True)
        ax.patch.set_visible(False)
        axes.append(ax)

    # ── Colors & markers ──
    palette = sns.color_palette("muted", n_colors=len(ys))
    markers = ['o','s','D','^','v','P','*','X','d']

    # ── Plot each series ──
    handles = []
    for i, (ax, y, lbl) in enumerate(zip(axes, ys, labels)):
        color = palette[i]
        m = markers[i % len(markers)]
        line, = ax.plot(
            x, y,
            color=color,
            marker=m, markersize=6,
            markerfacecolor='white', markeredgewidth=1.5,
            linewidth=2.2, alpha=0.9,
            label=lbl
        )
        handles.append(line)
        # style spine & ticks & label
        side = 'left' if i == 0 else 'right'
        ax.spines[side].set_color(color)
        ax.tick_params(axis='y', colors=color, labelsize=10)
        ax.yaxis.label.set_color(color)
        ax.set_ylabel(lbl, labelpad=15)
        # annotation with offset
        yoff = 8 if i % 2 == 0 else -10
        ax.annotate(
            f"{y.iloc[-1]:.2f}",
            xy=(x.iloc[-1], y.iloc[-1]),
            xytext=(5, yoff),
            textcoords="offset points",
            color=color,
            fontsize=9,
            fontweight='bold'
        )

    # ── Minor ticks & grid ──
    host.minorticks_on()
    host.grid(which='major', linestyle='-', linewidth=0.8, alpha=0.7)
    host.grid(which='minor', linestyle=':', linewidth=0.5, alpha=0.4)

    # ── Labels & title ──
    host.set_xlabel(x_label.get(), labelpad=15, fontsize=11)
    fig.suptitle(title_entry.get(), fontsize=16, fontweight='bold', y=0.96)

    # ── Adjust for legend ──
    fig.subplots_adjust(top=0.82, right=0.85)

    # ── Legend inside at top center, small font ──
    host.legend(
        handles, labels,
        loc='upper center',
        bbox_to_anchor=(0.5, 1.02),
        ncol=len(ys),
        fontsize='small',
        frameon=False
    )

    # ── Embed into Tkinter ──
    if canvas_plot:
        canvas_plot.get_tk_widget().destroy()
    canvas_plot = FigureCanvasTkAgg(fig, master=plot_area)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)

def save_plot():

    if not fig:
        return messagebox.showwarning("Nothing to save", "Plot first.")
    path = filedialog.asksaveasfilename(defaultextension=".png",
                                        filetypes=[("PNG files","*.png")])
    if path:
        fig.savefig(path)
        messagebox.showinfo("Saved", path)

def publish_gallery():
    if not saved_plots:
        return messagebox.showwarning("Empty Gallery", "Nothing has been added yet.")
    path = filedialog.asksaveasfilename(defaultextension=".png",
                                        filetypes=[("PNG files","*.png")])
    if not path:
        return

    n, cols = len(saved_plots), 2
    rows = math.ceil(n/cols)
    fig = plt.Figure(figsize=(cols*6, rows*4), dpi=200, tight_layout=True)
    axes = fig.subplots(nrows=rows, ncols=cols, squeeze=False)
    pal = sns.color_palette("deep", n_colors=2)

    for idx, cfg in enumerate(saved_plots):
        r, c = divmod(idx, cols)
        ax = axes[r][c]
        x  = cfg['x']
        ys = cfg['ys_series']
        ls = cfg['labels']

        if len(ys) == 1:
            # Single series: no markers, thin line
            ax.plot(x, ys[0],
                    label=ls[0],
                    color=pal[0],
                    linewidth=1.0,    # thinner line
                    alpha=0.9)        # slight opacity for detail
            ax.set_ylabel(ls[0], color=pal[0])
            ax.legend(fontsize='x-small', loc='upper right', frameon=False)

        elif len(ys) == 2:
            # First series on left axis
            ax.plot(x, ys[0],
                    label=ls[0],
                    color=pal[0],
                    linewidth=1.0,
                    alpha=0.9)
            ax.set_ylabel(ls[0], color=pal[0])
            # Second series on right axis
            ax2 = ax.twinx()
            ax2.plot(x, ys[1],
                     label=ls[1],
                     color=pal[1],
                     linewidth=1.0,
                     alpha=0.9)
            ax2.set_ylabel(ls[1], color=pal[1])
            # combined legend
            h1, l1 = ax.get_legend_handles_labels()
            h2, l2 = ax2.get_legend_handles_labels()
            ax.legend(h1+h2, l1+l2,
                      fontsize='x-small',
                      loc='upper right',
                      frameon=False)

        else:
            # fallback: plot all on one axis
            for i, (y, lbl) in enumerate(zip(ys, ls)):
                ax.plot(x, y,
                        label=lbl,
                        linewidth=1.0,
                        alpha=0.9)
            ax.legend(fontsize='x-small', loc='upper right', frameon=False)

        ax.set_title(cfg['title'], fontsize=10)
        ax.set_xlabel(cfg['x_label'], fontsize=9)
        ax.grid(alpha=0.3)

    # remove any empty subplots
    for idx in range(n, rows*cols):
        r, c = divmod(idx, cols)
        fig.delaxes(axes[r][c])

    fig.savefig(path)
    messagebox.showinfo("Gallery Saved", f"Saved to:\n{path}")
    

def edit_gallery():
    if not saved_plots:
        return messagebox.showinfo("Gallery Empty", "No saved plots to edit.")

    # New pop-up
    win = tk.Toplevel(root)
    win.title("Edit Gallery")
    win.geometry("400x300")

    # Listbox showing saved plot titles
    lb = tk.Listbox(win, font=("Arial", 10), activestyle='none')
    lb.pack(side='left', fill='both', expand=True, padx=(10,0), pady=10)

    # Populate it
    def refresh_list():
        lb.delete(0, tk.END)
        for i, cfg in enumerate(saved_plots):
            lb.insert(tk.END, f"{i+1}. {cfg['title']}")

    refresh_list()

    # Scrollbar for the listbox
    sb = ttk.Scrollbar(win, orient='vertical', command=lb.yview)
    lb.configure(yscrollcommand=sb.set)
    sb.pack(side='left', fill='y', pady=10)

    # Right-hand frame for buttons
    btns = ttk.Frame(win)
    btns.pack(side='right', fill='y', padx=10, pady=10)

    def delete_selected():
        sel = lb.curselection()
        if not sel:
            return messagebox.showwarning("No selection", "Please select a plot to delete.")
        idx = sel[0]
        # remove it
        del saved_plots[idx]
        refresh_list()

    ttk.Button(btns, text="Delete Selected", command=delete_selected) \
        .pack(fill='x', pady=(0,5))
    ttk.Button(btns, text="Close", command=win.destroy) \
        .pack(fill='x')
    
# — UI setup —
root = tk.Tk()
root.title("Flexible Multi‑Y‑Axis Plotter")
root.geometry("1000x700")

style = ttk.Style(root)
style.theme_use('clam')
style.configure("My.TLabelframe.Label", font=("Arial",11,"bold"))

# PanedWindow for adjustable divider
paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
paned.pack(fill='both', expand=True)

# Sidebar (left)
sidebar = ttk.Frame(paned, width=320)
sidebar.pack_propagate(False)
paned.add(sidebar, weight=0)

# Plot area (right)
plot_area = ttk.Frame(paned)
paned.add(plot_area, weight=1)

# Scrollable sidebar
sidebar_canvas = tk.Canvas(sidebar, borderwidth=0)
sidebar_scroll = ttk.Scrollbar(sidebar, orient='vertical', command=sidebar_canvas.yview)
sidebar_canvas.configure(yscrollcommand=sidebar_scroll.set)
sidebar_canvas.pack(side='left', fill='both', expand=True)
sidebar_scroll.pack(side='right', fill='y')

scrollable_frame = ttk.Frame(sidebar_canvas)
window_id = sidebar_canvas.create_window((0,0), window=scrollable_frame, anchor='nw')
scrollable_frame.bind("<Configure>", lambda e: sidebar_canvas.configure(scrollregion=sidebar_canvas.bbox("all")))
sidebar_canvas.bind("<Configure>", lambda e: sidebar_canvas.itemconfig(window_id, width=e.width))

# --- Controls in scrollable_frame ---

ttk.Label(scrollable_frame, text="Data Sources", style="My.TLabelframe.Label")\
    .pack(anchor='w', pady=(10,0))
ttk.Button(scrollable_frame, text="Load Excel/CSV…", command=load_files)\
    .pack(fill='x', pady=5)

ttk.Separator(scrollable_frame, orient='horizontal').pack(fill='x', pady=5)

ttk.Label(scrollable_frame, text="X‑Axis Configuration", style="My.TLabelframe.Label")\
    .pack(anchor='w', pady=(10,0))

# X‑Axis: File
frm = ttk.Frame(scrollable_frame)
frm.pack(fill='x', pady=2)
ttk.Label(frm, text="File:").pack(side='left')
file_x_cb = ttk.Combobox(frm, state='readonly')
file_x_cb.pack(side='left', fill='x', expand=True)

# X‑Axis: Sheet
frm = ttk.Frame(scrollable_frame)
frm.pack(fill='x', pady=2)
ttk.Label(frm, text="Sheet:").pack(side='left')
sheet_x_cb = ttk.Combobox(frm, state='readonly')
sheet_x_cb.pack(side='left', fill='x', expand=True)

# X‑Axis: Column
frm = ttk.Frame(scrollable_frame)
frm.pack(fill='x', pady=2)
ttk.Label(frm, text="Column:").pack(side='left')
col_x_cb = ttk.Combobox(frm, state='readonly')
col_x_cb.pack(side='left', fill='x', expand=True)

h_f, h_s = make_file_handler(file_x_cb, sheet_x_cb, [col_x_cb])
file_x_cb.bind("<<ComboboxSelected>>", h_f)
sheet_x_cb.bind("<<ComboboxSelected>>", h_s)

ttk.Separator(scrollable_frame, orient='horizontal').pack(fill='x', pady=5)

header = ttk.Frame(scrollable_frame)
header.pack(fill='x', pady=(10,0))
ttk.Label(header, text="Y‑Axes", style="My.TLabelframe.Label").pack(side='left')
ttk.Button(header, text="+ Add Y‑Axis", command=add_y_axis).pack(side='right')

y_container = ttk.Frame(scrollable_frame)
y_container.pack(fill='x', pady=5)

ttk.Separator(scrollable_frame, orient='horizontal').pack(fill='x', pady=5)

ttk.Label(scrollable_frame, text="Plot Labels", style="My.TLabelframe.Label")\
    .pack(anchor='w', pady=(10,0))
ttk.Label(scrollable_frame, text="X‑Label:").pack(anchor='w', pady=2)
x_label = ttk.Entry(scrollable_frame); x_label.pack(fill='x', pady=2)
ttk.Label(scrollable_frame, text="Title:").pack(anchor='w', pady=2)
title_entry = ttk.Entry(scrollable_frame); title_entry.pack(fill='x', pady=2)

ttk.Button(scrollable_frame, text="Plot", command=plot).pack(fill='x', pady=10)
ttk.Button(scrollable_frame, text="Save Plot", command=save_plot).pack(fill='x')

ttk.Button(scrollable_frame, text="Add to Gallery",  command=save_current_plot_config) \
    .pack(fill='x', pady=5)
ttk.Button(scrollable_frame, text="Publish Gallery", command=publish_gallery) \
    .pack(fill='x', pady=5)

ttk.Button(scrollable_frame, text="Edit Gallery", command=edit_gallery) \
    .pack(fill='x', pady=5)

# Start the app
root.mainloop()

 
