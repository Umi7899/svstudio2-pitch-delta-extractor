var SCRIPT_TITLE = "Extract Rendered Pitch to Pitch Delta";

var SCOPE_SELECTED_FIRST = 0;
var SCOPE_ALL_NOTES = 1;
var SCOPE_SELECTED_ONLY = 2;

var PADDING_AUTO_1_16 = 0;
var PADDING_STRICT = 1;
var PADDING_CONNECTED_PHRASE = 2;

var SAMPLING_PROFILES = [
  { key: "balanced", sampleSeconds: 0.005, simplifyThreshold: 0.002 },
  { key: "high", sampleSeconds: 0.002, simplifyThreshold: 0.001 },
  { key: "fast", sampleSeconds: 0.010, simplifyThreshold: 0.01 }
];

var POLL_INTERVAL_MS = 200;
var MAX_WAIT_MS = 8000;
var PITCH_DELTA_MIN = -1200;
var PITCH_DELTA_MAX = 1200;

function getClientInfo() {
  return {
    "name": SV.T(SCRIPT_TITLE),
    "author": "Codex",
    "versionNumber": 1,
    "minEditorVersion": 0x020101
  };
}

function getTranslations(langCode) {
  if (langCode === "zh-cn") {
    return [
      [SCRIPT_TITLE, "提取渲染音高到音高偏差"],
      ["Scope", "作用范围"],
      ["Selected Notes First (fallback to all notes)", "选中音符优先（无选中则全部）"],
      ["All Notes in Current Group", "当前组全部音符"],
      ["Selected Notes Only", "仅选中音符"],
      ["Sampling Profile", "采样档位"],
      ["Balanced (5ms, low simplify)", "平衡（5ms，低强度简化）"],
      ["High Precision (2ms, minimal simplify)", "高精度（2ms，最小简化）"],
      ["High Performance (10ms, moderate simplify)", "高性能（10ms，中等简化）"],
      ["Padding Mode", "边界策略"],
      ["Auto extend by 1/16 quarter", "自动外扩 1/16 拍"],
      ["Strict selected boundaries", "严格按选择边界"],
      ["Connected phrase ranges", "扩展到连通乐句"],
      ["Pure AI baseline (clear existing pitchDelta first)", "纯 AI 基线（先清空既有音高偏差）"],
      ["No editable notes found in current group.", "当前组没有可处理的音符。"],
      ["Please select notes first in \"Selected Notes Only\" mode.", "“仅选中音符”模式下请先选中音符。"],
      ["Target group is referenced %d times in this project. Continue? This will affect all references to the same group.", "该目标音符组在工程中被引用 %d 次。继续将影响该组的所有引用，是否继续？"],
      ["No valid target range after applying scope and padding.", "应用范围和边界策略后没有可处理区间。"],
      ["No voiced pitch frame detected in rendered result.", "渲染结果未检测到有效有声音高帧。"],
      ["Waiting for computed pitch timed out.", "等待计算音高超时。"],
      ["Extraction complete.\nRanges: %d\nSampled frames: %d\nVisible points: %d\nClamped frames: %d\nElapsed: %.2fs", "提取完成。\n区间数：%d\n采样帧数：%d\n可见控制点：%d\n限幅帧数：%d\n耗时：%.2f 秒"],
      ["Extraction failed and has been rolled back.\nReason: %s", "提取失败，已回滚。\n原因：%s"]
    ];
  }
  return [];
}

function main() {
  var editor = SV.getMainEditor();
  var scope = editor.getCurrentGroup();
  if (!scope) {
    SV.showMessageBox(SV.T(SCRIPT_TITLE), SV.T("No editable notes found in current group."));
    SV.finish();
    return;
  }

  var group = scope.getTarget();
  if (!group || group.getNumNotes() <= 0) {
    SV.showMessageBox(SV.T(SCRIPT_TITLE), SV.T("No editable notes found in current group."));
    SV.finish();
    return;
  }

  var dialogResult = SV.showCustomDialog(buildDialogDefinition());
  if (!dialogResult.status) {
    SV.finish();
    return;
  }

  var options = normalizeDialogAnswers(dialogResult.answers);
  var selectedNotes = getSelectedNotesInGroup(editor.getSelection().getSelectedNotes(), group);
  var allNotes = getAllNotes(group);
  var targetNotes = chooseTargetNotes(selectedNotes, allNotes, options.scopeMode);
  if (targetNotes.length === 0) {
    SV.showMessageBox(
      SV.T(SCRIPT_TITLE),
      options.scopeMode === SCOPE_SELECTED_ONLY
        ? SV.T("Please select notes first in \"Selected Notes Only\" mode.")
        : SV.T("No editable notes found in current group."));
    SV.finish();
    return;
  }

  var usageCount = countGroupReferences(SV.getProject(), group);
  if (usageCount > 1) {
    var warning = formatMessage(
      SV.T("Target group is referenced %d times in this project. Continue? This will affect all references to the same group."),
      [usageCount]
    );
    if (!SV.showOkCancelBox(SV.T(SCRIPT_TITLE), warning)) {
      SV.finish();
      return;
    }
  }

  var ranges = buildTargetRanges(targetNotes, allNotes, options.paddingMode);
  if (ranges.length === 0) {
    SV.showMessageBox(SV.T(SCRIPT_TITLE), SV.T("No valid target range after applying scope and padding."));
    SV.finish();
    return;
  }

  var state = {
    editor: editor,
    scope: scope,
    group: group,
    groupRef: scope,
    groupRefOffset: scope.getTimeOffset(),
    groupPitchOffset: scope.getPitchOffset(),
    timeAxis: SV.getProject().getTimeAxis(),
    project: SV.getProject(),
    allNotes: allNotes,
    pitchDelta: group.getParameter("pitchDelta"),
    ranges: ranges,
    pureAiBaseline: options.pureAiBaseline,
    profile: SAMPLING_PROFILES[options.samplingProfile],
    startTimeMs: new Date().getTime(),
    backupPointsByRange: [],
    writtenPoints: 0,
    voicedFrames: 0,
    clampedPoints: 0
  };

  runExtraction(state);
}

function buildDialogDefinition() {
  return {
    "title": SV.T(SCRIPT_TITLE),
    "message": "",
    "buttons": "OkCancel",
    "widgets": [
      {
        "name": "scopeMode",
        "type": "ComboBox",
        "label": SV.T("Scope"),
        "choices": [
          SV.T("Selected Notes First (fallback to all notes)"),
          SV.T("All Notes in Current Group"),
          SV.T("Selected Notes Only")
        ],
        "default": SCOPE_SELECTED_FIRST
      },
      {
        "name": "samplingProfile",
        "type": "ComboBox",
        "label": SV.T("Sampling Profile"),
        "choices": [
          SV.T("Balanced (5ms, low simplify)"),
          SV.T("High Precision (2ms, minimal simplify)"),
          SV.T("High Performance (10ms, moderate simplify)")
        ],
        "default": 0
      },
      {
        "name": "paddingMode",
        "type": "ComboBox",
        "label": SV.T("Padding Mode"),
        "choices": [
          SV.T("Auto extend by 1/16 quarter"),
          SV.T("Strict selected boundaries"),
          SV.T("Connected phrase ranges")
        ],
        "default": PADDING_AUTO_1_16
      },
      {
        "name": "pureAiBaseline",
        "type": "CheckBox",
        "text": SV.T("Pure AI baseline (clear existing pitchDelta first)"),
        "default": true
      }
    ]
  };
}

function normalizeDialogAnswers(answers) {
  var profile = answers.samplingProfile;
  if (profile < 0 || profile >= SAMPLING_PROFILES.length) {
    profile = 0;
  }
  return {
    scopeMode: answers.scopeMode,
    samplingProfile: profile,
    paddingMode: answers.paddingMode,
    pureAiBaseline: true
  };
}

function getSelectedNotesInGroup(selectedNotes, group) {
  var notes = [];
  var targetId = getObjectId(group);
  for (var i = 0; i < selectedNotes.length; i++) {
    var note = selectedNotes[i];
    var parent = note.getParent();
    if (parent && getObjectId(parent) === targetId) {
      notes.push(note);
    }
  }
  return sortNotesByOnset(notes);
}

function getAllNotes(group) {
  var notes = [];
  for (var i = 0; i < group.getNumNotes(); i++) {
    notes.push(group.getNote(i));
  }
  return sortNotesByOnset(notes);
}

function chooseTargetNotes(selectedNotes, allNotes, scopeMode) {
  if (scopeMode === SCOPE_SELECTED_ONLY) {
    return selectedNotes.slice();
  }
  if (scopeMode === SCOPE_ALL_NOTES) {
    return allNotes.slice();
  }
  if (selectedNotes.length > 0) {
    return selectedNotes.slice();
  }
  return allNotes.slice();
}

function sortNotesByOnset(notes) {
  notes.sort(function(a, b) {
    var onsetDiff = a.getOnset() - b.getOnset();
    if (onsetDiff !== 0) {
      return onsetDiff;
    }
    return a.getEnd() - b.getEnd();
  });
  return notes;
}

function buildTargetRanges(targetNotes, allNotes, paddingMode) {
  if (targetNotes.length === 0 || allNotes.length === 0) {
    return [];
  }

  var baseRanges;
  if (paddingMode === PADDING_CONNECTED_PHRASE) {
    baseRanges = expandToConnectedPhraseRanges(targetNotes, allNotes);
  } else {
    baseRanges = mergeRanges(notesToRanges(targetNotes));
  }

  var minBlick = allNotes[0].getOnset();
  var maxBlick = allNotes[allNotes.length - 1].getEnd();
  var paddedRanges = [];
  var padding = paddingMode === PADDING_AUTO_1_16 ? (SV.QUARTER / 16) : 0;

  for (var i = 0; i < baseRanges.length; i++) {
    var start = baseRanges[i][0] - padding;
    var end = baseRanges[i][1] + padding;
    if (paddingMode !== PADDING_AUTO_1_16) {
      start = baseRanges[i][0];
      end = baseRanges[i][1];
    }

    start = Math.max(minBlick, Math.floor(start));
    end = Math.min(maxBlick, Math.ceil(end));
    if (end > start) {
      paddedRanges.push([start, end]);
    }
  }

  return mergeRanges(paddedRanges);
}

function notesToRanges(notes) {
  var ranges = [];
  for (var i = 0; i < notes.length; i++) {
    ranges.push([notes[i].getOnset(), notes[i].getEnd()]);
  }
  return ranges;
}

function mergeRanges(ranges) {
  if (ranges.length === 0) {
    return [];
  }
  ranges.sort(function(a, b) {
    return a[0] - b[0];
  });
  var merged = [[ranges[0][0], ranges[0][1]]];
  for (var i = 1; i < ranges.length; i++) {
    var curr = ranges[i];
    var last = merged[merged.length - 1];
    if (curr[0] <= last[1]) {
      if (curr[1] > last[1]) {
        last[1] = curr[1];
      }
    } else {
      merged.push([curr[0], curr[1]]);
    }
  }
  return merged;
}

function expandToConnectedPhraseRanges(targetNotes, allNotes) {
  var phrases = [];
  var phraseStart = allNotes[0].getOnset();
  var phraseEnd = allNotes[0].getEnd();

  for (var i = 1; i < allNotes.length; i++) {
    var note = allNotes[i];
    if (note.getOnset() <= phraseEnd) {
      if (note.getEnd() > phraseEnd) {
        phraseEnd = note.getEnd();
      }
    } else {
      phrases.push([phraseStart, phraseEnd]);
      phraseStart = note.getOnset();
      phraseEnd = note.getEnd();
    }
  }
  phrases.push([phraseStart, phraseEnd]);

  var targetRanges = notesToRanges(targetNotes);
  var expanded = [];
  for (var p = 0; p < phrases.length; p++) {
    for (var r = 0; r < targetRanges.length; r++) {
      if (rangesOverlap(phrases[p], targetRanges[r])) {
        expanded.push([phrases[p][0], phrases[p][1]]);
        break;
      }
    }
  }
  return mergeRanges(expanded);
}

function rangesOverlap(a, b) {
  return a[0] < b[1] && b[0] < a[1];
}

function countGroupReferences(project, targetGroup) {
  var targetId = getObjectId(targetGroup);
  var count = 0;
  for (var i = 0; i < project.getNumTracks(); i++) {
    var track = project.getTrack(i);
    for (var j = 0; j < track.getNumGroups(); j++) {
      var ref = track.getGroupReference(j);
      if (getObjectId(ref.getTarget()) === targetId) {
        count++;
      }
    }
  }
  return count;
}

function runExtraction(state) {
  backupPitchDelta(state);
  state.project.newUndoRecord();

  if (state.pureAiBaseline) {
    clearPitchDeltaInRanges(state);
  }

  processRangeSequentially(state, 0);
}

function backupPitchDelta(state) {
  state.backupPointsByRange = [];
  for (var i = 0; i < state.ranges.length; i++) {
    var start = state.ranges[i][0];
    var end = state.ranges[i][1];
    var points = state.pitchDelta.getPoints(start, end);
    state.backupPointsByRange.push({
      start: start,
      end: end,
      points: clonePoints(points)
    });
  }
}

function clearPitchDeltaInRanges(state) {
  for (var i = 0; i < state.ranges.length; i++) {
    state.pitchDelta.remove(state.ranges[i][0], state.ranges[i][1]);
  }
}

function clonePoints(points) {
  var cloned = [];
  for (var i = 0; i < points.length; i++) {
    cloned.push([points[i][0], points[i][1]]);
  }
  return cloned;
}

function processRangeSequentially(state, rangeIndex) {
  if (rangeIndex >= state.ranges.length) {
    finalizeSuccess(state);
    return;
  }

  var range = state.ranges[rangeIndex];
  sampleRangeAndWritePitchDelta(state, range, function(err, result) {
    if (err) {
      rollbackAndFail(state, err.message || String(err));
      return;
    }

    state.writtenPoints += result.pointsWritten;
    state.voicedFrames += result.voicedFrames;
    state.clampedPoints += result.clampedPoints;
    if (state.profile.simplifyThreshold > 0) {
      state.pitchDelta.simplify(range[0], range[1], state.profile.simplifyThreshold);
    }
    processRangeSequentially(state, rangeIndex + 1);
  });
}

function sampleRangeAndWritePitchDelta(state, range, done) {
  var absoluteStart = range[0] + state.groupRefOffset;
  var absoluteEnd = range[1] + state.groupRefOffset;
  var sampleIntervalBlick = getSamplingIntervalBlick(state.timeAxis, absoluteStart, state.profile.sampleSeconds);
  var frameCount = Math.floor((absoluteEnd - absoluteStart) / sampleIntervalBlick) + 1;
  if (frameCount <= 0) {
    done(null, { pointsWritten: 0, voicedFrames: 0 });
    return;
  }

  pollComputedPitch(
    state.groupRef,
    absoluteStart,
    sampleIntervalBlick,
    frameCount,
    POLL_INTERVAL_MS,
    MAX_WAIT_MS,
    function(err, sampledPitch) {
      if (err) {
        // If rendered pitch is unavailable (e.g. trial voicebank limits or deleted lyrics),
        // fill this range with zero anchors instead of failing the whole script.
        if (isTimeoutError(err)) {
          var zeroFallback = writePitchDeltaFromSampledPitch(
            state,
            range,
            absoluteStart,
            sampleIntervalBlick,
            []
          );
          done(null, zeroFallback);
          return;
        }
        done(err);
        return;
      }
      var writeResult = writePitchDeltaFromSampledPitch(
        state,
        range,
        absoluteStart,
        sampleIntervalBlick,
        sampledPitch
      );
      if (writeResult.voicedFrames <= 0) {
        var zeroFallback = writePitchDeltaFromSampledPitch(
          state,
          range,
          absoluteStart,
          sampleIntervalBlick,
          []
        );
        done(null, zeroFallback);
        return;
      }
      done(null, writeResult);
    }
  );
}

function getSamplingIntervalBlick(timeAxis, absoluteStartBlick, sampleSeconds) {
  var startSec = timeAxis.getSecondsFromBlick(absoluteStartBlick);
  var nextBlick = timeAxis.getBlickFromSeconds(startSec + sampleSeconds);
  var delta = Math.round(nextBlick - absoluteStartBlick);
  if (delta < 1) {
    delta = 1;
  }
  return delta;
}

function pollComputedPitch(groupRef, absStart, intervalBlick, frameCount, pollMs, maxWaitMs, callback) {
  var startedMs = new Date().getTime();

  function tryFetch() {
    var pitches = SV.getComputedPitchForGroup(
      groupRef,
      absStart,
      intervalBlick,
      frameCount
    );
    if (pitches && pitches.length > 0) {
      callback(null, pitches);
      return;
    }

    var elapsed = new Date().getTime() - startedMs;
    if (elapsed >= maxWaitMs) {
      callback(new Error(SV.T("Waiting for computed pitch timed out.")));
      return;
    }

    SV.setTimeout(pollMs, tryFetch);
  }

  tryFetch();
}

function isTimeoutError(err) {
  if (!err) {
    return false;
  }
  return String(err.message || err) === SV.T("Waiting for computed pitch timed out.");
}

function writePitchDeltaFromSampledPitch(state, range, absoluteStart, sampleIntervalBlick, sampledPitch) {
  var pointsWritten = 0;
  var voicedFrames = 0;
  var clampedPoints = 0;
  var writtenBlickMap = {};
  var noteEdgeMap = {};
  var rangeNotes = getNotesIntersectingRange(state.allNotes, range);

  function addPoint(blick, value) {
    var roundedBlick = Math.round(blick);
    state.pitchDelta.add(roundedBlick, value);
    var key = String(roundedBlick);
    if (!writtenBlickMap[key]) {
      writtenBlickMap[key] = true;
      pointsWritten++;
    }
  }

  var noteIndex = findNoteIndexAtOrBefore(state.allNotes, range[0]);
  if (noteIndex < 0) {
    noteIndex = 0;
  }

  for (var i = 0; i < sampledPitch.length; i++) {
    var rawPitch = sampledPitch[i];
    if (rawPitch === null || rawPitch === undefined || !isFinite(rawPitch) || rawPitch <= 0) {
      continue;
    }
    var pitchSemitone = rawPitch;

    var absoluteBlick = absoluteStart + i * sampleIntervalBlick;
    var localBlick = Math.round(absoluteBlick - state.groupRefOffset);
    if (localBlick < range[0] || localBlick >= range[1]) {
      continue;
    }

    while (noteIndex < state.allNotes.length && state.allNotes[noteIndex].getEnd() <= localBlick) {
      noteIndex++;
    }
    if (noteIndex >= state.allNotes.length) {
      break;
    }

    var note = state.allNotes[noteIndex];
    if (!(note.getOnset() <= localBlick && localBlick < note.getEnd())) {
      continue;
    }

    voicedFrames++;

    var detuneSemitone = 0;
    if (typeof note.getDetune === "function") {
      detuneSemitone = note.getDetune() / 100.0;
    }

    var noteBase = note.getPitch() + detuneSemitone;
    var deltaNoRefOffset = (pitchSemitone - noteBase) * 100.0;
    var deltaWithRefOffset = (pitchSemitone - (noteBase + state.groupPitchOffset)) * 100.0;
    var delta = chooseReasonableDelta(deltaNoRefOffset, deltaWithRefOffset, state.groupPitchOffset);
    if (delta > PITCH_DELTA_MAX) {
      delta = PITCH_DELTA_MAX;
      clampedPoints++;
    } else if (delta < PITCH_DELTA_MIN) {
      delta = PITCH_DELTA_MIN;
      clampedPoints++;
    }

    addPoint(localBlick, delta);

    var noteKey = getNoteKey(note);
    if (!noteEdgeMap[noteKey]) {
      noteEdgeMap[noteKey] = {
        firstBlick: localBlick,
        firstDelta: delta,
        lastBlick: localBlick,
        lastDelta: delta
      };
    } else {
      if (localBlick < noteEdgeMap[noteKey].firstBlick) {
        noteEdgeMap[noteKey].firstBlick = localBlick;
        noteEdgeMap[noteKey].firstDelta = delta;
      }
      if (localBlick > noteEdgeMap[noteKey].lastBlick) {
        noteEdgeMap[noteKey].lastBlick = localBlick;
        noteEdgeMap[noteKey].lastDelta = delta;
      }
    }
  }

  // Anchor each note head/tail to prevent interpolation carryover from adjacent rests.
  for (var n = 0; n < rangeNotes.length; n++) {
    var rangeNote = rangeNotes[n];
    var noteStart = Math.max(range[0], rangeNote.getOnset());
    var noteEndExclusive = Math.min(range[1], rangeNote.getEnd());
    if (noteEndExclusive <= noteStart) {
      continue;
    }
    var noteEndInclusive = noteEndExclusive - 1;
    var edge = noteEdgeMap[getNoteKey(rangeNote)];
    if (edge) {
      addPoint(noteStart, edge.firstDelta);
      addPoint(noteEndInclusive, edge.lastDelta);
    } else {
      addPoint(noteStart, 0);
      addPoint(noteEndInclusive, 0);
    }
  }

  // Force pitchDelta to 0 in non-note regions so disconnected notes do not bridge.
  addRestZeroAnchors(range, rangeNotes, addPoint);

  return {
    pointsWritten: pointsWritten,
    voicedFrames: voicedFrames,
    clampedPoints: clampedPoints
  };
}

function chooseReasonableDelta(deltaNoRefOffset, deltaWithRefOffset, groupPitchOffset) {
  if (!groupPitchOffset) {
    return deltaNoRefOffset;
  }
  if (Math.abs(deltaWithRefOffset) < Math.abs(deltaNoRefOffset)) {
    return deltaWithRefOffset;
  }
  return deltaNoRefOffset;
}

function findNoteIndexAtOrBefore(notes, blick) {
  if (notes.length === 0) {
    return -1;
  }
  var lo = 0;
  var hi = notes.length - 1;
  var ans = -1;
  while (lo <= hi) {
    var mid = Math.floor((lo + hi) / 2);
    if (notes[mid].getOnset() <= blick) {
      ans = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return ans;
}

function getNotesIntersectingRange(notes, range) {
  var ret = [];
  for (var i = 0; i < notes.length; i++) {
    if (notes[i].getEnd() <= range[0]) {
      continue;
    }
    if (notes[i].getOnset() >= range[1]) {
      break;
    }
    ret.push(notes[i]);
  }
  return ret;
}

function addRestZeroAnchors(range, rangeNotes, addPoint) {
  var cursor = range[0];
  for (var i = 0; i < rangeNotes.length; i++) {
    var noteStart = Math.max(rangeNotes[i].getOnset(), range[0]);
    var noteEnd = Math.min(rangeNotes[i].getEnd(), range[1]);
    if (noteEnd <= noteStart) {
      continue;
    }
    if (noteStart > cursor) {
      addZeroSegment(cursor, noteStart, addPoint);
    }
    if (noteEnd > cursor) {
      cursor = noteEnd;
    }
  }
  if (cursor < range[1]) {
    addZeroSegment(cursor, range[1], addPoint);
  }
}

function addZeroSegment(startInclusive, endExclusive, addPoint) {
  if (endExclusive <= startInclusive) {
    return;
  }
  var left = Math.round(startInclusive);
  var right = Math.round(endExclusive - 1);
  if (right < left) {
    right = left;
  }
  addPoint(left, 0);
  if (right !== left) {
    addPoint(right, 0);
  }
}

function getNoteKey(note) {
  if (typeof note.getIndexInParent === "function") {
    return "idx:" + String(note.getIndexInParent());
  }
  return [
    String(note.getOnset()),
    String(note.getEnd()),
    String(note.getPitch())
  ].join("|");
}

function rollbackAndFail(state, reason) {
  for (var i = 0; i < state.ranges.length; i++) {
    state.pitchDelta.remove(state.ranges[i][0], state.ranges[i][1]);
  }
  for (var j = 0; j < state.backupPointsByRange.length; j++) {
    var backup = state.backupPointsByRange[j];
    for (var k = 0; k < backup.points.length; k++) {
      state.pitchDelta.add(backup.points[k][0], backup.points[k][1]);
    }
  }

  SV.showMessageBox(
    SV.T(SCRIPT_TITLE),
    formatMessage(SV.T("Extraction failed and has been rolled back.\nReason: %s"), [reason])
  );
  SV.finish();
}

function finalizeSuccess(state) {
  var elapsed = (new Date().getTime() - state.startTimeMs) / 1000.0;
  var visiblePoints = countPointsInRanges(state.pitchDelta, state.ranges);
  SV.showMessageBox(
    SV.T(SCRIPT_TITLE),
    formatMessage(
      SV.T("Extraction complete.\nRanges: %d\nSampled frames: %d\nVisible points: %d\nClamped frames: %d\nElapsed: %.2fs"),
      [state.ranges.length, state.writtenPoints, visiblePoints, state.clampedPoints, elapsed]
    )
  );
  SV.finish();
}

function countPointsInRanges(automation, ranges) {
  var total = 0;
  for (var i = 0; i < ranges.length; i++) {
    total += automation.getPoints(ranges[i][0], ranges[i][1]).length;
  }
  return total;
}

function formatMessage(template, args) {
  var out = template;
  for (var i = 0; i < args.length; i++) {
    out = out.replace(/%(\.\d+)?[dfs]/, function(token, precision) {
      var value = args[i];
      if (token[token.length - 1] === "d") {
        return String(Math.round(value));
      }
      if (token[token.length - 1] === "f") {
        var digits = 6;
        if (precision) {
          digits = parseInt(precision.substring(1), 10);
        }
        return Number(value).toFixed(digits);
      }
      return String(value);
    });
  }
  return out;
}

function getObjectId(obj) {
  if (!obj) {
    return "";
  }
  if (typeof obj.getUUID === "function") {
    return obj.getUUID();
  }
  return String(obj);
}
