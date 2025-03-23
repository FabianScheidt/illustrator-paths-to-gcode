// Offset of the pen relative to the origin
var PEN_OFFSET_X = 35;
var PEN_OFFSET_Y = 23;

// Z height and feed when drawing
var DRAW_HEIGHT = 0.3;
var DRAW_FEED = 2000;

// Z height and feed when moving between paths
var LIFT_HEIGHT = DRAW_HEIGHT + 2;
var LIFT_FEED = 6000;

// Height for putting in the paper
var HIGH_HEIGHT = 30;

// Filter for the spot color name (null will disable the filter)
var SPOT_COLOR_FILTER = "Plotter";

// Conversion between Illustrator's unit (points) and G-Code (millimeter)
var POINTS_TO_MILLIMETER = 0.352777777777;

// Max error when interpolating Bézier curves into line segments
var MAX_ERROR_MILLIMETER = 0.1;
var MAX_ERROR_POINTS = MAX_ERROR_MILLIMETER / POINTS_TO_MILLIMETER;

// G-Code executed before and after the actual paths
var GCODE_BEFORE_ALL = [
    ";FLAVOR:Marlin",
    "G28 ; Home all axes",
    "G21 ; Set units to mm",
    "G90 ; Absolute positioning",
];
var GCODE_BEFORE_ARTBOARD = [
    "G0 F" + LIFT_FEED.toFixed(3) + " X0 Y220 Z" + HIGH_HEIGHT.toFixed(3) + " ; Make room to put in the paper",
    "M0 ; Wait to put in paper",
    "M75 ; Start timer",
];
var GCODE_AFTER_ARTBOARD = [
    "M75 ; Pause timer",
];
var GCODE_AFTER_ALL = [
    "M77 ; Stop timer",
    "G0 X0.0 Y220.0 Z" + HIGH_HEIGHT.toFixed(3) + " ; Present result",
    "M84 X Y ; Disable all steppers but Z",
];

/**
 * Runs the conversion script
 */
function run() {
    if (app.documents.length === 0) {
        alert("No document open!");
        return;
    }
    var doc = app.activeDocument;
    var pathItems = filterPathItemsBySpotColor(doc.pathItems);
    var gCode = [];

    // Create G-Code for each individual artboard
    for (var i = 0; i < doc.artboards.length; i++) {
        var artboard = doc.artboards[i];
        var artboardPathItems = filterPathItemsByArtboard(pathItems, artboard);
        gCode = gCode.concat(
            ["; Start of Artboard \"" + artboard.name + "\""],
            GCODE_BEFORE_ARTBOARD,
            [""],
            convertPathsToGCode(artboardPathItems, artboard),
            GCODE_AFTER_ARTBOARD,
            ["; End of Artboard \"" + artboard.name + "\"", ""],
        )
    }

    var file = File.saveDialog("Select a location to save the G-Code output", "*.gcode");
    saveGCodeFile(gCode, file);
    alert("G-Code exported to " + file.fsName + "!");
}

/**
 * Filter the list of path items to only those that match the spot color filter (if set)
 * @param pathItems
 * @returns {*[]}
 */
function filterPathItemsBySpotColor(pathItems) {
    if (SPOT_COLOR_FILTER === null) {
        return pathItems;
    }

    var res = [];
    for (var i = 0; i < pathItems.length; i++) {
        var pathItem = pathItems[i];
        var strokeColor = pathItem.strokeColor;
        if (strokeColor.typename !== "SpotColor") {
            continue;
        }
        if (strokeColor.spot.name.toLowerCase().indexOf(SPOT_COLOR_FILTER.toLowerCase()) > -1) {
            res.push(pathItem);
        }
    }
    return res;
}

/**
 * Filter the list of path items to only those that are on the specified artboard
 * @param pathItems
 * @param artboard
 * @returns {*[]}
 */
function filterPathItemsByArtboard(pathItems, artboard) {
    var res = [];
    for (var i = 0; i < pathItems.length; i++) {
        var point = pathItems[i].position;

        var left = artboard.artboardRect[0];
        var top = artboard.artboardRect[1];
        var right = artboard.artboardRect[2];
        var bottom = artboard.artboardRect[3];

        if (point[0] >= left && point[0] <= right && point[1] <= top && point[1] >= bottom) {
            res.push(pathItems[i]);
        }
    }
    return res;
}

/**
 * Converts a list of path items into a list of G-Code instructions.
 * @param pathItems
 * @param artboard
 * @returns {string[]}
 */
function convertPathsToGCode(pathItems,  artboard) {
    var gcode = [];
    for (var i = 0; i < pathItems.length; i++) {
        var pathItem = pathItems[i];

        if (pathItem.pathPoints.length === 0) {
            continue;
        }

        var flattenedPoints = flattenPathPoints(pathItem, 0.05);

        var firstPoint = mapCoordinates(flattenedPoints[0], artboard);
        gcode.push("G0 F" + LIFT_FEED + " X" + firstPoint[0].toFixed(3) + " Y" + firstPoint[1].toFixed(3) + " ; Move to start of path");
        gcode.push("G0 Z" + DRAW_HEIGHT.toFixed(3) + " ; Lower pen");

        for (var j = 1; j < flattenedPoints.length; j++) {
            var point = mapCoordinates(flattenedPoints[j], artboard);
            var addFeed = j === 1 ? " F" + DRAW_FEED : "";
            gcode.push("G1" + addFeed + " X" + point[0].toFixed(3) + " Y" + point[1].toFixed(3));
        }

        gcode.push("G0 Z" + LIFT_HEIGHT.toFixed(3) + " ; Lift pen");
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
 * @param artboard
 * @returns {number[]}
 */
function mapCoordinates(point, artboard) {
    return [
        // X-axis matches the printer.
        (point[0] - artboard.artboardRect[0]) * POINTS_TO_MILLIMETER + PEN_OFFSET_X,
        // Y-axis is inverted and has it's origin at the top left of the document, unlike the printer at the bottom right.
        (point[1] - artboard.artboardRect[3]) * POINTS_TO_MILLIMETER + PEN_OFFSET_Y,
    ];
}

/**
 * Saves G-Code to a file, adding the configured code before and after.
 * @param gCode {string[]}
 * @param file {File}
 */
function saveGCodeFile(gCode, file) {
    var gcodeStr = [].concat(
        GCODE_BEFORE_ALL,
        [""],
        gCode,
        GCODE_AFTER_ALL,
        [""],
    ).join("\n");

    file.encoding = "utf-8";
    file.open("w");
    file.write(gcodeStr);
    file.close();
}

// All methods declared. Run!
run();
