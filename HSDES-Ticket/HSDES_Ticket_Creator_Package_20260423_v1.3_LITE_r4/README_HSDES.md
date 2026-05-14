# HSDES Ticket Creator（终端用户步骤）

## 1. 解压并进入目录

```powershell
cd <解压后的目录>
```

## 2. 安装依赖

```powershell
pip install -r requirements.txt
```

## 3. 填写模板

1. 打开 `WO_Template.xlsx`。
2. 每行填写一条数据，至少要有 `Title`。
3. `ticket_id` 列规则：
   - 有值：按该 ID 对比并更新原 ticket（仅更新差异字段）。
   - 空值：新建 ticket。
4. 新增模板列默认会自动识别：
  - 已在 `field_mapping` 中配置的列，仍按映射后的 API 字段发送。
  - 未配置映射的新列，会按“Excel 列名 = API 字段名”自动透传。
  - `ticket_id` 属于控制列，不会发到 HSDES。
  - 如果某个新列不想自动透传，可在配置文件里设置 `"auto_map_unmapped_fields": false`。
5. 保存并关闭 Excel 文件（必须关闭，避免文件占用）。

## 4. 先预览（可选）

```powershell
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --dry-run
```

## 5. 正式执行

```powershell
python hsdes_ticket_creator.py --excel "WO_Template.xlsx"
```

## 6. 看结果

1. 终端会打印每行对应的 ticket_id。
2. 脚本会输出 `hsdes_create_results.json`。
3. 脚本会把 ticket_id 回写到 `WO_Template.xlsx`。
4. 预检查状态在 `hsdes_precheck_status.json`：
   - `ok: true` 表示文件可写并已开始执行。
   - `ok: false` 表示文件被占用或不可写。

## 7. 常用命令（复制即用）

```powershell
# 指定工作表
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --sheet "WO_Template"

# 从第 N 条开始（N 从 0 开始）
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --start-row 0
```

从指定行继续：

```powershell
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --start-row 5
```

## 输出文件

运行结束后会生成 hsdes_create_results.json，记录每行的 payload、创建结果和返回的 ticket ID。

如果接口返回的 ticket ID 不在 id、ticket_id、data.id、data.ticket_id 这些常见字段中，需要按真实返回结构调整 hsdes_ticket_creator.py 里的 extract_ticket_id。

## Cookie 缓存与登录问题

### Cookie 缓存位置

默认保存在当前系统用户目录下：`%USERPROFILE%\.hsdes_cookie_cache`

```powershell
# 查看当前缓存的Cookie（仅所有者可读）
Get-Content "$env:USERPROFILE\.hsdes_cookie_cache"

# 手动清除缓存（下次执行会重新登录）
Remove-Item "$env:USERPROFILE\.hsdes_cookie_cache" -ErrorAction SilentlyContinue

# 自定义缓存文件路径（适合团队共享脚本但隔离缓存）
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --cookie-cache-file "C:\Users\<you>\.hsdes_cache"

# 完全禁用缓存（仅本次运行）
python hsdes_ticket_creator.py --excel "WO_Template.xlsx" --no-cookie-cache
```

### 常见问题

**Q: 为什么每次都需要登录？**
- A: Cookie 可能已过期或无效。脚本会自动检测并提示重新登录。如需强制清除缓存，执行 `Remove-Item "$env:USERPROFILE\.hsdes_cookie_cache" -ErrorAction SilentlyContinue`

**Q: Cookie 缓存多久有效？**
- A: 取决于 HSDES 服务端的 session 超时设置，通常为几小时到几天。脚本在每次执行时会自动验证 Cookie 是否仍然有效。

**Q: 如何手动提供 Cookie？**
- A: 从浏览器 DevTools 的 Network 或 Application 标签复制 Cookie，然后：
  ```powershell
  $env:HSDES_COOKIE = Get-Clipboard
  python hsdes_ticket_creator.py --excel "WO_Template.xlsx"
  ```

**Q: 脚本能否保存密码？**
- A: 不能。出于安全考虑，脚本不会保存密码。但可以通过以下方式避免每次输入：
  - 方式1：使用命令行参数 `--username ... --password ...`（只有脚本执行时可见）
  - 方式2：设置环境变量，然后在PowerShell脚本中使用
  - 方式3：使用 Cookie 缓存机制（推荐）

## 故障排查

**认证失败：**
```
❌ 登录失败: 账户或密码错误 (HTTP 401)
```
- 检查账户名和密码是否正确
- 确认 Intel 账户在 HSDES 中是否有权限

**连接失败：**
```
❌ 登录异常: ...
```
- 检查网络连接和是否在 VPN 上
- 确认 HSDES URL 是否可访问

**Cookie 验证失败：**
```
✗ 缓存的Cookie已过期或无效
```
- Cookie 已过期，脚本会自动提示重新登录
- 也可以手动执行 `Remove-Item "$env:USERPROFILE\.hsdes_cookie_cache" -ErrorAction SilentlyContinue` 清除缓存
