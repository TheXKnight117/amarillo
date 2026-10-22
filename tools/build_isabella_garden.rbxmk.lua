local input, output, mosaicPath = ...
assert(input and output and mosaicPath, "usage: build_isabella_garden.rbxmk.lua <input.rbxl> <output.rbxl> <mosaics.json>")

local game = fs.read(input, "rbxl")
local workspace = game:GetService("Workspace")
local mosaics = fs.read(mosaicPath, "json").photos

local old = workspace:FindFirstChild("Para Isabella")
if old then
	old.Parent = nil
end

local root = Instance.new("Model", workspace)
root.Name = "Para Isabella"

local pink = Color3.fromRGB(255, 90, 165)
local lightPink = Color3.fromRGB(255, 194, 222)
local blush = Color3.fromRGB(255, 225, 238)
local white = Color3.fromRGB(255, 255, 255)
local gold = Color3.fromRGB(255, 214, 112)
local plum = Color3.fromRGB(82, 31, 73)

local function makePart(name, size, cframe, color, material, canCollide, parent)
	local part = Instance.new("Part", parent or root)
	part.Name = name
	part.Anchored = true
	part.CanCollide = canCollide == true
	part.CanTouch = false
	part.CanQuery = false
	part.CastShadow = false
	part.Size = size
	part.CFrame = cframe
	part.Color = color
	part.Material = material or Enum.Material.SmoothPlastic
	part.TopSurface = Enum.SurfaceType.Smooth
	part.BottomSurface = Enum.SurfaceType.Smooth
	return part
end

local function addSurfaceText(part, face, text, color, background, textSize)
	local gui = Instance.new("SurfaceGui", part)
	gui.Name = "Dedicatoria"
	gui.Face = face
	gui.CanvasSize = Vector2.new(900, 420)
	gui.SizingMode = Enum.SurfaceGuiSizingMode.FixedSize
	gui.LightInfluence = 0
	gui.AlwaysOnTop = false

	local label = Instance.new("TextLabel", gui)
	label.Name = "Texto"
	label.Size = UDim2.fromScale(1, 1)
	label.Position = UDim2.fromScale(0, 0)
	label.BackgroundColor3 = background
	label.BackgroundTransparency = 0.12
	label.BorderSizePixel = 0
	label.Text = text
	label.TextColor3 = color
	label.TextScaled = true
	label.TextWrapped = true
	label.TextSize = textSize or 48
	label.TextStrokeColor3 = plum
	label.TextStrokeTransparency = 0.55
	label.Font = Enum.Font.GothamBold
	return gui
end

local function addBillboardHeart(position, scale)
	local anchor = makePart(
		"Corazón flotante",
		Vector3.new(0.2, 0.2, 0.2),
		CFrame.new(position),
		pink,
		Enum.Material.Neon,
		false
	)
	anchor.Transparency = 1

	local gui = Instance.new("BillboardGui", anchor)
	gui.Name = "Corazón"
	gui.Size = UDim2.fromOffset(90 * scale, 90 * scale)
	gui.AlwaysOnTop = true
	gui.LightInfluence = 0

	local label = Instance.new("TextLabel", gui)
	label.Name = "Símbolo"
	label.Size = UDim2.fromScale(1, 1)
	label.BackgroundTransparency = 1
	label.Text = "♥"
	label.TextColor3 = pink
	label.TextScaled = true
	label.TextStrokeColor3 = white
	label.TextStrokeTransparency = 0.15
	label.Font = Enum.Font.GothamBlack
end

local function addPhotoDisplay(photo, index, cframe, caption)
	local model = Instance.new("Model", root)
	model.Name = string.format("Recuerdo %d - %s", index, caption)

	local width = 20
	local height = 15
	local board = makePart(
		"Foto mosaico",
		Vector3.new(width, height, 0.65),
		cframe,
		plum,
		Enum.Material.SmoothPlastic,
		false,
		model
	)

	local glow = makePart(
		"Resplandor",
		Vector3.new(width + 1.5, height + 1.5, 0.35),
		cframe * CFrame.new(0, 0, 0.18),
		pink,
		Enum.Material.Neon,
		false,
		model
	)
	glow.Transparency = 0.18

	local canvasWidth = 480
	local canvasHeight = 360
	local pixelWidth = canvasWidth / photo.width
	local pixelHeight = canvasHeight / photo.height

	local gui = Instance.new("SurfaceGui", board)
	gui.Name = "Foto integrada"
	gui.Face = Enum.NormalId.Front
	gui.CanvasSize = Vector2.new(canvasWidth, canvasHeight)
	gui.SizingMode = Enum.SurfaceGuiSizingMode.FixedSize
	gui.LightInfluence = 0
	gui.AlwaysOnTop = false

	for runIndex, run in ipairs(photo.runs) do
		local pixel = Instance.new("Frame", gui)
		pixel.Name = tostring(runIndex)
		pixel.BorderSizePixel = 0
		pixel.BackgroundColor3 = Color3.fromRGB(run[4], run[5], run[6])
		pixel.Position = UDim2.fromOffset(run[1] * pixelWidth, run[2] * pixelHeight)
		pixel.Size = UDim2.fromOffset(run[3] * pixelWidth, pixelHeight)
	end

	local top = makePart(
		"Marco superior",
		Vector3.new(width + 2, 1, 1),
		cframe * CFrame.new(0, height / 2 + 0.5, 0),
		gold,
		Enum.Material.Neon,
		false,
		model
	)
	local bottom = makePart(
		"Marco inferior",
		Vector3.new(width + 2, 1, 1),
		cframe * CFrame.new(0, -height / 2 - 0.5, 0),
		gold,
		Enum.Material.Neon,
		false,
		model
	)
	local left = makePart(
		"Marco izquierdo",
		Vector3.new(1, height, 1),
		cframe * CFrame.new(-width / 2 - 0.5, 0, 0),
		gold,
		Enum.Material.Neon,
		false,
		model
	)
	local right = makePart(
		"Marco derecho",
		Vector3.new(1, height, 1),
		cframe * CFrame.new(width / 2 + 0.5, 0, 0),
		gold,
		Enum.Material.Neon,
		false,
		model
	)
	top.Transparency = 0.05
	bottom.Transparency = 0.05
	left.Transparency = 0.05
	right.Transparency = 0.05

	local plaqueCFrame = cframe * CFrame.new(0, -height / 2 - 2.1, 0)
	local plaque = makePart(
		"Mensaje",
		Vector3.new(14, 2.2, 0.8),
		plaqueCFrame,
		lightPink,
		Enum.Material.Neon,
		false,
		model
	)
	addSurfaceText(plaque, Enum.NormalId.Front, caption .. "  ♥", plum, blush, 40)

	local light = Instance.new("PointLight", board)
	light.Name = "Luz romántica"
	light.Color = lightPink
	light.Brightness = 1.25
	light.Range = 24
	light.Shadows = false
end

-- Entrance arch, placed just after the existing spawn point.
makePart(
	"Pilar izquierdo",
	Vector3.new(1.2, 14, 1.2),
	CFrame.new(27, 49, -469),
	gold,
	Enum.Material.Neon,
	false
)
makePart(
	"Pilar derecho",
	Vector3.new(1.2, 14, 1.2),
	CFrame.new(27, 49, -441),
	gold,
	Enum.Material.Neon,
	false
)
local entrance = makePart(
	"Bienvenida",
	Vector3.new(1, 7, 30),
	CFrame.new(27, 59, -455),
	lightPink,
	Enum.Material.Neon,
	false
)
addSurfaceText(
	entrance,
	Enum.NormalId.Left,
	"PARA ISABELLA  ♥\nUn jardín para la persona que hace florecer mis días",
	plum,
	blush,
	50
)
addSurfaceText(
	entrance,
	Enum.NormalId.Right,
	"PARA ISABELLA  ♥\nUn jardín para la persona que hace florecer mis días",
	plum,
	blush,
	50
)

-- Three memories alternate along the path toward the torii gate.
addPhotoDisplay(mosaics[1], 1, CFrame.new(55, 54, -431), "Tú y yo")
addPhotoDisplay(mosaics[2], 2, CFrame.new(95, 51, -486) * CFrame.Angles(0, math.pi, 0), "Nuestro mundo juntos")
addPhotoDisplay(mosaics[3], 3, CFrame.new(132, 48, -429), "Siempre a tu lado")

local heartPositions = {
	Vector3.new(38, 51, -470),
	Vector3.new(72, 50, -441),
	Vector3.new(112, 48, -474),
	Vector3.new(151, 46, -438),
	Vector3.new(181, 45, -473),
	Vector3.new(212, 46, -439),
	Vector3.new(242, 47, -473),
}
for index, position in ipairs(heartPositions) do
	addBillboardHeart(position, index % 2 == 0 and 0.85 or 1)
end

-- Large glowing heart and final dedication near the shrine.
local heartCenter = Vector3.new(270, 55, -455)
local heartPieces = 34
for index = 0, heartPieces - 1 do
	local angle = (index / heartPieces) * math.pi * 2
	local horizontal = 16 * math.sin(angle) ^ 3
	local vertical =
		13 * math.cos(angle)
		- 5 * math.cos(2 * angle)
		- 2 * math.cos(3 * angle)
		- math.cos(4 * angle)
	local color = index % 2 == 0 and pink or lightPink
	local orb = makePart(
		"Corazón final",
		Vector3.new(2.8, 2.8, 2.8),
		CFrame.new(heartCenter.X, heartCenter.Y + vertical * 0.72, heartCenter.Z + horizontal * 0.72),
		color,
		Enum.Material.Neon,
		false
	)
	orb.Shape = Enum.PartType.Ball
end

local finalSign = makePart(
	"Dedicatoria final",
	Vector3.new(0.8, 10, 20),
	CFrame.new(269.5, 55, -455),
	plum,
	Enum.Material.SmoothPlastic,
	false
)
addSurfaceText(
	finalSign,
	Enum.NormalId.Left,
	"ISABELLA\nCada momento contigo es mi lugar favorito.\nGracias por hacer mis días más bonitos.\nTE AMO  ♥",
	white,
	plum,
	46
)

local finalLight = Instance.new("PointLight", finalSign)
finalLight.Name = "Luz final"
finalLight.Color = pink
finalLight.Brightness = 2
finalLight.Range = 36
finalLight.Shadows = true

fs.write(output, game, "rbxl")

print("Created romantic garden for Isabella")
print("Photo mosaics:", #mosaics)
print("Output:", output)
