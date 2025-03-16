// Offset of the pen relative to the origin
var PEN_OFFSET_X = 35
var PEN_OFFSET_Y = 23

// Z height and feed when drawing
var DRAW_HEIGHT = 0;
var DRAW_FEED = 2000;

// Z height and feed when moving between paths
var LIFT_HEIGHT = DRAW_HEIGHT + 2;
var LIFT_FEED = 6000;

// Height for putting in the paper
var HIGH_HEIGHT = 50;

// Conversion between Illustrator's unit (points) and G-Code (millimeter)
var POINTS_TO_MILLITMETER = 0.352777777777;

// Max error when interpolating Bézier curves into line segments
var MAX_ERROR_MILLIMETER = 0.1;
var MAX_ERROR_POINTS = MAX_ERROR_MILLIMETER / POINTS_TO_MILLITMETER;

// G-Code executed before and after the actual paths
var GCODE_BEFORE = [
    ";FLAVOR:Marlin",
    "G28 ; Home all axes",
    "G21 ; Set units to mm",
    "G90 ; Absolute positioning",
    "G0 F" + LIFT_FEED + " X0 Y220 Z" + HIGH_HEIGHT + " ; Make room to put in the paper",
    "M0 ; Wait to put in paper",
    "G0 F" + LIFT_FEED + " X" + PEN_OFFSET_X + " Y" + PEN_OFFSET_Y + " Z" + LIFT_HEIGHT + "; Move to origin at lift height",
    "M75 ; Start timer",
    "",
];
var GCODE_AFTER = [
    "M77 ; Stop timer",
    "G0 X0.0 Y220.0 Z20.0 ; Present result",
    "M84 X Y ; Disable all steppers but Z",
    "",
];

/**
 * Runs the conversion script
 */
function run() {
    if (app.documents.length === 0) {
        alert("No document open!");
        return;
    }
    var pathItems = app.activeDocument.pathItems;
    var gCode = convertPathsToGCode(pathItems);

    var file = File.saveDialog("Select a location to save the G-Code output", "*.gcode");
    saveGCodeFile(gCode, file);
    alert("G-Code exported to " + file.fsName + "!");
}

/**
 * Converts a list of path items into a list of G-Code instructions.
 * @param pathItems
 * @returns {string[]}
 */
function convertPathsToGCode(pathItems) {
    var gcode = [];
    for (var i = 0; i < pathItems.length; i++) {
        var pathItem = pathItems[i];

        if (pathItem.pathPoints.length === 0) {
            continue;
        }

        var flattenedPoints = flattenPathPoints(pathItem, 0.05);

        var firstPoint = mapCoordinates(flattenedPoints[0]);
        gcode.push("G0 F" + LIFT_FEED + " X" + firstPoint[0].toFixed(2) + " Y" + firstPoint[1].toFixed(2) + " ; Move to start of path");
        gcode.push("G0 Z" + DRAW_HEIGHT.toFixed(2) + " ; Lower pen");

        for (var j = 1; j < flattenedPoints.length; j++) {
            var point = mapCoordinates(flattenedPoints[j]);
            var addFeed = j === 1 ? " F" + DRAW_FEED : "";
            gcode.push("G1" + addFeed + " X" + point[0].toFixed(2) + " Y" + point[1].toFixed(2));
        }

        gcode.push("G0 Z" + LIFT_HEIGHT.toFixed(2) + " ; Lift pen");
        gcode.push("");
    }
    return gcode;
}

/**
 * Reads points from a path item, interpolates Bézier curves and turns it into a flat list of points.
 * @param pathItem
 * @returns {number[][]}
 */
function flattenPathPoints(pathItem) {
    // Create a list from our path points. Close the path if needed.
    var pathPoints = [];
    for (var i = 0; i < pathItem.pathPoints.length; i++) {
        pathPoints.push(pathItem.pathPoints[i]);
    }
    if (pathItem.closed) {
        pathPoints.push(pathItem.pathPoints[0]);
    }

    // Now iterate list turn bezier curves into line segments
    var points = [pathPoints[0].anchor];
    for (var j = 1; j < pathPoints.length; j++) {
        var pA = pathPoints[j - 1];
        var pB = pathPoints[j];

        points = points.concat(
            subdivideBezierToLineSegments(pA.anchor, pA.rightDirection, pB.leftDirection, pB.anchor, MAX_ERROR_POINTS, [])
        );
    }

    return points;
}

/**
 * Returns a list of points that approximate a Bézier curve as line segments, making sure that a maximum error is not exceeded.
 * @param p0 {number[]}
 * @param p1 {number[]}
 * @param p2 {number[]}
 * @param p3 {number[]}
 * @param maxError {number}
 * @param points {number[][]}
 * @returns {number[][]}
 */
function subdivideBezierToLineSegments(p0, p1, p2, p3, maxError, points) {
    if (!points) points = [p0];

    // Compute the Bézier midpoint (De Casteljau's algorithm)
    var p01 = midpoint(p0, p1);
    var p12 = midpoint(p1, p2);
    var p23 = midpoint(p2, p3);
    var p012 = midpoint(p01, p12);
    var p123 = midpoint(p12, p23);
    var bezMid = midpoint(p012, p123);

    // Compute the midpoint of the straight line and the deviation error
    var lineMid = midpoint(p0, p3);
    var error = distance(bezMid, lineMid);

    if (error <= maxError) {
        // If the error is within tolerance, approximate with a straight line
        points.push(p3);
    } else {
        // Recursively subdivide
        subdivideBezierToLineSegments(p0, p01, p012, bezMid, maxError, points);
        subdivideBezierToLineSegments(bezMid, p123, p23, p3, maxError, points);
    }

    return points;
}

/**
 * Calculates the midpoint between two points
 * @param pA {number[]}
 * @param pB {number[]}
 * @returns {number[]}
 */
function midpoint(pA, pB) {
    return [
        (pA[0] + pB[0]) / 2,
        (pA[1] + pB[1]) / 2,
    ];
}

/**
 * Euclidean distance between two points
 * @param pA {number[]}
 * @param pB {number[]}
 * @returns {number}
 */
function distance(pA, pB) {
    dx = pA[0] - pB[0];
    dy = pA[1] - pB[1];
    return Math.sqrt(dx * dx + dy * dy);
}

/**
 * Maps a point from Illustrator's coordinate system the coordinate system of the 3D printer
 * @param point {number[]}
 * @returns {number[]}
 */
function mapCoordinates(point) {
    // Determine artboard height
    var doc = app.activeDocument;
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var artboardHeight = (artboard[1] - artboard[3]) * POINTS_TO_MILLITMETER;

    // Convert Illustrator's points to mm and flip y-axis
    return [
        // X-axis matches the printer.
        point[0] * POINTS_TO_MILLITMETER + PEN_OFFSET_X,
        // Y-axis is inverted and has it's origin at the top left of the document, unlike the printer at the bottom right.
        point[1] * POINTS_TO_MILLITMETER + PEN_OFFSET_Y + artboardHeight,
    ];
}

/**
 * Saves G-Code to a file, adding the configured code before and after.
 * @param gCode {string[]}
 * @param file {File}
 */
function saveGCodeFile(gCode, file) {
    var gcodeStr = [].concat(
        GCODE_BEFORE,
        gCode,
        GCODE_AFTER,
    ).join("\n");

    file.open("w");
    file.write(gcodeStr);
    file.close();
}

// All methods declared. Run!
run();
