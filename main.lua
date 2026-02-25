--- @since 25.2.7

local M = {}

function M:peek(job)
	local start, cache = os.clock(), ya.file_cache(job)
	if not cache then
		return
	end

	local ok, err = self:preload(job)
	if not ok or err then
		return
	end

	ya.sleep(math.max(0, rt.preview.image_delay / 1000 + start - os.clock()))
	ya.image_show(cache, job.area)
	ya.preview_widgets(job, {})
end

function M:seek(job)
	local h = cx.active.current.hovered
	if h and h.url == job.file.url then
		local step = ya.clamp(-1, job.units, 1)
		ya.manager_emit("peek", { math.max(0, cx.active.preview.skip + step), only_if = job.file.url })
	end
end

local function tmp_base()
	if ya.target_family() == "windows" then
		return os.getenv("TEMP") or os.getenv("TMP") or "."
	end
	return os.getenv("TMPDIR") or "/tmp"
end

function M:doc2pdf(job)
	local st = fs.cha(job.file.url)
	if not st then
		return nil, Err("Failed to stat `%s`", tostring(job.file.url))
	end

	-- Deterministic cache key so we can reuse conversions
	local key_src = table.concat({
		"office.yazi",
		tostring(job.file.url),
		tostring(st.mtime or ""),
		tostring(st.len or ""),
		tostring(job.skip + 1),
	}, "|")

	local key = ya.hash(key_src)
	local tmp = tmp_base() .. "/yazi-office/" .. key .. "/"

	local ok, err = fs.create("dir_all", Url(tmp))
	if not ok then
		return nil, Err("Failed to create temp dir %s: %s", tmp, err)
	end

	local pdf = tmp .. job.file.name:gsub("%.[^%.]+$", (".p" .. tostring(job.skip + 1) .. ".pdf"))

	-- If already converted, reuse it
	local f = io.open(pdf, "rb")
	if f then
		f:close()
		return pdf
	end

	local args = {
		"--headless",
		"--convert-to",
		('pdf:draw_pdf_Export:{"PageRange":{"type":"string","value":"%d"}}'):format(job.skip + 1),
		"--outdir",
		tmp,
		tostring(job.file.url),
	}

	local out = Command("soffice")
			:arg(args)
			:stdin(Command.NULL)
			:stdout(Command.PIPED)
			:stderr(Command.PIPED)
			:output()

	if not out then
		return nil, Err("Failed to start `soffice` (even though it works in PowerShell). Check PATH visibility for Yazi.")
	end

	if not out.status.success then
		local output = (out.stdout or "") .. (out.stderr or "")
		return nil, Err("soffice failed converting `%s`: %s", job.file.name, output)
	end

	-- LibreOffice outputs <originalname>.pdf, rename to include page number so pages don't collide
	local produced = tmp .. job.file.name:gsub("%.[^%.]+$", ".pdf")
	local ok_mv, mv_err = fs.rename(Url(produced), Url(pdf))
	if not ok_mv then
		-- If rename fails (sometimes output already correct), just try reading produced
		local fp = io.open(produced, "rb")
		if fp then
			fp:close()
			return produced
		end
		return nil, Err("Failed to finalize PDF. rename error: %s", mv_err)
	end

	return pdf
end

function M:preload(job)
	local cache = ya.file_cache(job)
	if not cache or fs.cha(cache) then
		return true
	end

	local tmp_pdf, err = self:doc2pdf(job)
	if not tmp_pdf then
		return true, Err("    " .. "%s", err)
	end

	local output, err = Command("pdftoppm")
			:arg({
				"-singlefile",
				"-jpeg",
				"-jpegopt",
				"quality=" .. rt.preview.image_quality,
				"-f",
				1,
				tostring(tmp_pdf),
			})
			:stdout(Command.PIPED)
			:stderr(Command.PIPED)
			:output()

	local rm_tmp_pdf, rm_err = fs.remove("file", Url(tmp_pdf))
	if not rm_tmp_pdf then
		return true, Err("Failed to remove %s, error: %s", tmp_pdf, rm_err)
	end

	if not output then
		return true, Err("Failed to start `pdftoppm`, error: %s", err)
	elseif not output.status.success then
		local pages = tonumber(output.stderr:match("the last page %((%d+)%)")) or 0
		if job.skip > 0 and pages > 0 then
			ya.mgr_emit("peek", { math.max(0, pages - 1), only_if = job.file.url, upper_bound = true })
		end
		return true, Err("Failed to convert %s to image, stderr: %s", tmp_pdf, output.stderr)
	end

	return fs.write(cache, output.stdout)
end

return M
