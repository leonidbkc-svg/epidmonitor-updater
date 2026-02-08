def reset_canvas(current_canvas):
    if current_canvas.get("obj"):
        current_canvas["obj"].get_tk_widget().destroy()
        current_canvas["obj"] = None


def barh_with_value_labels(
    ax,
    labels,
    values,
    *,
    title,
    xlabel,
    colors=None,
    xlim_pad=0.15,
    value_fmt=None,
):
    bars = ax.barh(labels, values, color=colors)
    ax.set_title(title)
    ax.set_xlabel(xlabel)

    max_v = max(values) if values else 0
    ax.set_xlim(0, max_v * (1 + xlim_pad) if max_v else 10)

    for bar in bars:
        w = bar.get_width()
        text = value_fmt(w) if value_fmt else str(int(w))
        ax.text(
            w + (max_v * 0.01 + 0.5 if max_v else 0.5),
            bar.get_y() + bar.get_height() / 2,
            text,
            va="center",
            fontsize=9
        )


def bar_with_value_labels(
    ax,
    labels,
    values,
    *,
    title,
    colors=None,
    ylim_pad=0.2,
    value_fmt=None,
):
    bars = ax.bar(labels, values, color=colors)
    ax.set_title(title)

    max_v = max(values) if values else 0
    ax.set_ylim(0, max_v * (1 + ylim_pad) if max_v else 10)

    for bar in bars:
        h = bar.get_height()
        text = value_fmt(h) if value_fmt else str(int(h))
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h + (max_v * 0.03 if max_v else 0.3),
            text,
            ha="center",
            va="bottom",
            fontsize=9
        )


def stacked_barh(
    ax,
    group_names,
    loci_order,
    groups_items,
    *,
    colors_map,
    edgecolor="white",
    linewidth=0.6
):
    left = [0] * len(group_names)

    for locus in loci_order:
        values = []
        for items in groups_items:
            v = next((int(i["count"]) for i in items if i["name"] == locus), 0)
            values.append(v)

        if sum(values) == 0:
            continue

        ax.barh(
            group_names,
            values,
            left=left,
            label=locus,
            color=colors_map.get(locus, "#cccccc"),
            edgecolor=edgecolor,
            linewidth=linewidth
        )
        left = [l + v for l, v in zip(left, values)]
