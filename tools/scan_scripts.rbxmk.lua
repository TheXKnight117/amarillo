local input = ...
assert(input, "usage: scan_scripts.rbxmk.lua <place.rbxl>")

local game = fs.read(input, "rbxl")
local scriptClasses = {
	LocalScript = true,
	ModuleScript = true,
	Script = true,
}
local suspiciousPatterns = {
	"loadstring",
	"HttpGet",
	"HttpPost",
	"GetAsync",
	"PostAsync",
	"require%s*%(%s*%d+",
	"discord",
	"webhook",
}

local total = 0
local suspicious = 0
for _, instance in ipairs(game:GetDescendants()) do
	if scriptClasses[instance.ClassName] then
		total = total + 1
		local source = tostring(instance.Source or "")
		for _, pattern in ipairs(suspiciousPatterns) do
			if string.find(source, pattern) then
				suspicious = suspicious + 1
				print(string.format("SUSPICIOUS\t%s\t%s", pattern, instance:GetFullName()))
				break
			end
		end
	end
end

print(string.format("Scanned %d scripts; %d matched suspicious patterns", total, suspicious))
if suspicious > 0 then
	error("suspicious script patterns detected")
end
