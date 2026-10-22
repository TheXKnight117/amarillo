local input = ...
assert(input, "usage: inspect_spawns.rbxmk.lua <place.rbxl>")

local game = fs.read(input)
local workspace = game:GetService("Workspace")

local function property(instance, name)
	local ok, value = pcall(function()
		return instance[name]
	end)
	return ok and tostring(value) or "<unset>"
end

for _, instance in ipairs(workspace:GetDescendants()) do
	if instance.ClassName == "SpawnLocation" then
		local position = instance.CFrame.Position
		print(string.format(
			"Spawn %s pos=(%.2f, %.2f, %.2f) Enabled=%s Neutral=%s Transparency=%s",
			instance:GetFullName(),
			position.X,
			position.Y,
			position.Z,
			property(instance, "Enabled"),
			property(instance, "Neutral"),
			property(instance, "Transparency")
		))
	elseif instance.ClassName == "Camera" then
		local position = instance.CFrame.Position
		print(string.format(
			"Camera %s pos=(%.2f, %.2f, %.2f)",
			instance:GetFullName(),
			position.X,
			position.Y,
			position.Z
		))
	end
end
