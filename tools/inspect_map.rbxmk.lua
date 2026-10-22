local input = ...
assert(input, "usage: rbxmk run tools/inspect_map.rbxmk.lua -- <place.rbxl>")

local game = fs.read(input, "rbxl")
local workspace = game:GetService("Workspace")

local counts = {}
local parts = {}
local basePartClasses = {
	CornerWedgePart = true,
	MeshPart = true,
	Part = true,
	SpawnLocation = true,
	TrussPart = true,
	UnionOperation = true,
	VehicleSeat = true,
	WedgePart = true,
}

for _, instance in ipairs(workspace:GetDescendants()) do
	counts[instance.ClassName] = (counts[instance.ClassName] or 0) + 1
	if basePartClasses[instance.ClassName] then
		table.insert(parts, instance)
	end
end

local classes = {}
for className, count in pairs(counts) do
	table.insert(classes, { className, count })
end
table.sort(classes, function(a, b)
	return a[1] < b[1]
end)

print("Workspace:", workspace.Name)
print("Descendants:", #workspace:GetDescendants())
for _, item in ipairs(classes) do
	print(string.format("%-28s %d", item[1], item[2]))
end

print("\nBase parts:")
for _, part in ipairs(parts) do
	local position = part.CFrame.Position
	local lookVector = part.CFrame.LookVector
	local hasSize, size = pcall(function()
		return part.Size
	end)
	local sizeText = "unavailable"
	if hasSize then
		sizeText = string.format("%.2f, %.2f, %.2f", size.X, size.Y, size.Z)
	end
	print(string.format(
		"%s\t%s\tpos=(%.2f, %.2f, %.2f)\tlook=(%.2f, %.2f, %.2f)\tsize=(%s)",
		part.ClassName,
		part:GetFullName(),
		position.X,
		position.Y,
		position.Z,
		lookVector.X,
		lookVector.Y,
		lookVector.Z,
		sizeText
	))
end
