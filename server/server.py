"""stdio MCP server for OriginPro automation (plot + linear fit + PNG export).

Tools appear as mcp__origin__origin_plot after DSH registration. Requires Windows,
Python with pywin32 + originpro, and OriginPro installed."""
import json
import os
import sys
import traceback

LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lib")
if LIB not in sys.path:
    sys.path.insert(0, LIB)

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

# Windows 控制台默认 GBK：强制 UTF-8，否则中文报错文本会炸掉 JSON-RPC。
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def _out(sub=""):
    """输出目录：LAB_OUT_DIR > <cwd>/lab-out（不可写则逐级回退，跨工作区只读 cwd 不崩溃）。"""
    def _try(base):
        d = os.path.join(base, sub) if sub else base
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".write-test")
            open(probe, "w").close()
            os.remove(probe)
            return d
        except OSError:
            return None

    env = os.environ.get("LAB_OUT_DIR")
    for cand in ([env] if env else []) + [os.path.join(os.getcwd(), "lab-out"),
                                          os.path.expandvars(r"%LOCALAPPDATA%\lab-out"),
                                          os.path.join(os.path.expanduser("~"), "Desktop", "ds", "lab-out")]:
        d = _try(cand)
        if d:
            return d
    import tempfile
    return tempfile.mkdtemp(prefix="lab-out-")


def tool_origin_plot(args):
    """数据列 → Origin 出图 → PNG。args: {x:[], y:[], xlname?, ylname?, title?, fit?:bool}"""
    from origin_helper import ensure_origin
    o = ensure_origin()
    x = [float(v) for v in args["x"]]
    y = [float(v) for v in args["y"]]
    wb = o.new_book('w', 'mcp')
    ws = wb[0]
    ws.from_list(0, x, lname=args.get("xlname", "X"))
    ws.from_list(1, y, lname=args.get("ylname", "Y"))
    gp = o.find_graph("gmcp")
    if gp:
        gp.destroy()
    gp = o.new_graph("gmcp")
    gl = gp[0]
    gl.add_plot(ws, coly=1, colx=0)
    gl.rescale()
    png = os.path.join(_out("origin"), (args.get("title") or "plot") + ".png")
    gp.save_fig(png, type="png")
    result = {"png": png, "points": len(x)}
    if args.get("fit"):
        fit = o.LinearFit()
        fit.set_data(ws, 'A', 'B')
        res = fit.result()
        p = res.get('Parameters', {})
        rs = res.get('RegStats', {}).get('C1', {})
        result["fit"] = {
            "slope": p.get('Slope', {}).get('Value'),
            "intercept": p.get('Intercept', {}).get('Value'),
            "R2": rs.get('RSqCOD'),
        }
    return result



TOOL_SPECS = {
"origin_plot": {
        "description": ("Plot XY data in OriginPro and export PNG; optional linear fit "
                        "returning slope/intercept/R2. Args: x[], y[], xlname, ylname, title, fit(bool)."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "x": {"type": "array", "items": {"type": "number"}},
                "y": {"type": "array", "items": {"type": "number"}},
                "xlname": {"type": "string"}, "ylname": {"type": "string"},
                "title": {"type": "string"}, "fit": {"type": "boolean"},
            },
            "required": ["x", "y"],
        },
        "fn": tool_origin_plot,
    }
}


# ---------------------------------------------------------------- stdio ----
def _send(msg):
    body = json.dumps(msg, ensure_ascii=False)
    sys.stdout.write(body + "\n")
    sys.stdout.flush()


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue
        rid = req.get("id")
        method = req.get("method", "")
        params = req.get("params") or {}

        if method == "initialize":
            _send({"jsonrpc": "2.0", "id": rid, "result": {
                "protocolVersion": params.get("protocolVersion", "2024-11-05"),
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "origin", "version": "1.0.0"},
            }})
        elif method.startswith("notifications/"):
            continue
        elif method == "ping":
            _send({"jsonrpc": "2.0", "id": rid, "result": {}})
        elif method == "tools/list":
            tools = [{"name": n, "description": s["description"], "inputSchema": s["inputSchema"]}
                     for n, s in TOOL_SPECS.items()]
            _send({"jsonrpc": "2.0", "id": rid, "result": {"tools": tools}})
        elif method == "tools/call":
            name = params.get("name")
            spec = TOOL_SPECS.get(name)
            if spec is None:
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "isError": True, "content": [{"type": "text", "text": f"unknown tool {name}"}]}})
                continue
            try:
                out = spec["fn"](params.get("arguments") or {})
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "content": [{"type": "text", "text": json.dumps(out, ensure_ascii=False)}]}})
            except Exception as exc:
                tb = traceback.format_exc(limit=4)
                _send({"jsonrpc": "2.0", "id": rid, "result": {
                    "isError": True, "content": [{"type": "text", "text": f"{exc}\n{tb}"}]}})
        else:
            _send({"jsonrpc": "2.0", "id": rid,
                   "error": {"code": -32601, "message": f"method not found: {method}"}})


if __name__ == "__main__":
    main()
