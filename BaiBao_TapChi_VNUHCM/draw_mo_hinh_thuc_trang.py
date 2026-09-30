import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig, ax = plt.subplots(figsize=(10, 5))

# Define box coordinates (x, y, width, height)
# All boxes are on the same vertical scale, just spread out.
box_a = {'x': 0.05, 'y': 0.4, 'w': 0.22, 'h': 0.2, 'text': 'Chủ thể quyền\nSHTT'}
box_b = {'x': 0.4, 'y': 0.7, 'w': 0.22, 'h': 0.2, 'text': 'Nền tảng số /\nSàn TMĐT'}
box_c = {'x': 0.4, 'y': 0.1, 'w': 0.22, 'h': 0.2, 'text': 'Cơ quan Quản lý\nNhà nước'}
box_d = {'x': 0.75, 'y': 0.4, 'w': 0.22, 'h': 0.2, 'text': 'Tài khoản /\nCửa hàng vi phạm'}

def draw_box(ax, box):
    p = patches.Rectangle((box['x'], box['y']), box['w'], box['h'], 
                          fill=True, facecolor='white', edgecolor='black', lw=2, zorder=3)
    ax.add_patch(p)
    ax.text(box['x'] + box['w']/2, box['y'] + box['h']/2, box['text'], 
            ha='center', va='center', fontsize=11, fontweight='bold', zorder=4)

draw_box(ax, box_a)
draw_box(ax, box_b)
draw_box(ax, box_c)
draw_box(ax, box_d)

# Draw arrows (Folded lines - angle connecting)
# A -> B (Right then Up, or Up then Right)
ax.annotate("", xy=(box_b['x'], box_b['y'] + box_b['h']/2), 
            xytext=(box_a['x'] + box_a['w']/2, box_a['y'] + box_a['h']),
            arrowprops=dict(arrowstyle="->", color="black", lw=2, connectionstyle="angle,angleA=0,angleB=90,rad=0"), zorder=2)
ax.text(box_a['x'] + box_a['w']/2 + 0.01, box_b['y'] + box_b['h']/2 + 0.03, "1. Yêu cầu gỡ bỏ", fontsize=10, ha='left')

# A -> C (Right then Down, or Down then Right)
ax.annotate("", xy=(box_c['x'], box_c['y'] + box_c['h']/2), 
            xytext=(box_a['x'] + box_a['w']/2, box_a['y']),
            arrowprops=dict(arrowstyle="->", color="black", lw=2, connectionstyle="angle,angleA=0,angleB=-90,rad=0"), zorder=2)
ax.text(box_a['x'] + box_a['w']/2 + 0.01, box_c['y'] + box_c['h']/2 + 0.03, "3. Gửi tố cáo", fontsize=10, ha='left')

# B -> D (Right then Down)
ax.annotate("", xy=(box_d['x'] + box_d['w']/2, box_d['y'] + box_d['h']), 
            xytext=(box_b['x'] + box_b['w'], box_b['y'] + box_b['h']/2),
            arrowprops=dict(arrowstyle="->", color="black", lw=2, connectionstyle="angle,angleA=270,angleB=0,rad=0"), zorder=2)
ax.text(box_b['x'] + box_b['w'] + 0.01, box_b['y'] + box_b['h']/2 + 0.03, "2. Xử lý thụ động", fontsize=10, ha='left')

# C -> D (Right then Up)
ax.annotate("", xy=(box_d['x'] + box_d['w']/2, box_d['y']), 
            xytext=(box_c['x'] + box_c['w'], box_c['y'] + box_c['h']/2),
            arrowprops=dict(arrowstyle="->", color="black", lw=2, connectionstyle="angle,angleA=90,angleB=0,rad=0"), zorder=2)
ax.text(box_c['x'] + box_c['w'] + 0.01, box_c['y'] + box_c['h']/2 + 0.03, "4. Xử phạt", fontsize=10, ha='left')

ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis('off')
plt.savefig('mo_hinh_thuc_trang.png', dpi=300, bbox_inches='tight')
