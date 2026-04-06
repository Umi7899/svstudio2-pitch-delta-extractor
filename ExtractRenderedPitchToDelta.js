var SCRIPT_TITLE = "提取渲染音高到音高偏差";

var SCOPE_SELECTED_FIRST = 0;
var SCOPE_ALL_NOTES = 1;
var SCOPE_SELECTED_ONLY = 2;

var OUTPUT_PITCH_DELTA = 0;
var OUTPUT_PITCH_CONTROL_CURVE = 1;

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
var MANAGED_PITCH_CONTROL_KEY = "ExtractRenderedPitchToDelta.managedPitchControl";

function getClientInfo() {
  return {
    "name": SV.T(SCRIPT_TITLE),
    "author": "Codex",
    "versionNumber": 2,
    "minEditorVersion": 0x020101
  };
}

function getTranslations(langCode) {
  return [];
}

function main() {
  var state = null;
  try {
    var editor = SV.getMainEditor();
    var scope = editor.getCurrentGroup();
    if (!scope) {
      SV.showMessageBox(SCRIPT_TITLE, "当前组里没有可处理的音符。");
      SV.finish();
      return;
    }

    var group = scope.getTarget();
    if (!group || group.getNumNotes() <= 0) {
      SV.showMessageBox(SCRIPT_TITLE, "当前组里没有可处理的音符。");
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
        SCRIPT_TITLE,
        options.scopeMode === SCOPE_SELECTED_ONLY
          ? "“仅选中音符”模式下请先选中音符。"
          : "当前组里没有可处理的音符。");
      SV.finish();
      return;
    }

    var usageCount = countGroupReferences(SV.getProject(), group);
    if (usageCount > 1) {
      var warning = formatMessage(
        "该目标组在工程中被引用了 %d 次。继续执行会同时影响这些引用，是否继续？",
        [usageCount]
      );
      if (!SV.showOkCancelBox(SCRIPT_TITLE, warning)) {
        SV.finish();
        return;
      }
    }

    var ranges = buildTargetRanges(targetNotes, allNotes, options.paddingMode);
    if (ranges.length === 0) {
      SV.showMessageBox(SCRIPT_TITLE, "应用作用范围和边界策略后，没有可处理的有效区间。");
      SV.finish();
      return;
    }

    state = {
      editor: editor,
      scope: scope,
      group: group,
      groupRef: scope,
      groupRefTrackIndex: scope.getParent() ? scope.getParent().getIndexInParent() : -1,
      groupRefIndex: scope.getIndexInParent(),
      groupRefOffset: scope.getTimeOffset(),
      groupPitchOffset: scope.getPitchOffset(),
      timeAxis: SV.getProject().getTimeAxis(),
      project: SV.getProject(),
      allNotes: allNotes,
      pitchDelta: group.getParameter("pitchDelta"),
      ranges: ranges,
      outputMode: options.outputMode,
      pureAiBaseline: options.pureAiBaseline,
      profile: SAMPLING_PROFILES[options.samplingProfile],
      startTimeMs: new Date().getTime(),
      sessionId: String(new Date().getTime()),
      isFinalized: false,
      backupPointsByRange: [],
      backupPitchControls: [],
      backupNotePitchModes: [],
      writtenPoints: 0,
      outputObjectsWritten: 0,
      voicedFrames: 0,
      clampedPoints: 0
    };

    runExtraction(state);
  } catch (err) {
    handleFatalScriptError(err);
  }
}

function buildDialogDefinition() {
  return {
    "title": SCRIPT_TITLE,
    "message": "",
    "buttons": "OkCancel",
    "widgets": [
      {
        "name": "scopeMode",
        "type": "ComboBox",
        "label": "作用范围",
        "choices": [
          "选中音符优先（无选中则全部）",
          "当前组全部音符",
          "仅选中音符"
        ],
        "default": SCOPE_SELECTED_FIRST
      },
      {
        "name": "outputMode",
        "type": "ComboBox",
        "label": "输出模式",
        "choices": [
          "音高偏差（相对 cents）",
          "音高控制曲线（绝对音高覆盖）"
        ],
        "default": OUTPUT_PITCH_DELTA
      },
      {
        "name": "samplingProfile",
        "type": "ComboBox",
        "label": "采样档位",
        "choices": [
          "平衡（5ms，低强度简化）",
          "高精度（2ms，最小简化）",
          "高性能（10ms，中等简化）"
        ],
        "default": 0
      },
      {
        "name": "paddingMode",
        "type": "ComboBox",
        "label": "边界策略",
        "choices": [
          "自动外扩 1/16 拍",
          "严格按选择边界",
          "扩展到连通乐句"
        ],
        "default": PADDING_AUTO_1_16
      },
      {
        "name": "pureAiBaseline",
        "type": "CheckBox",
        "text": "写入前先替换目标区间内之前的提取结果",
        "default": true
      }
    ]
  };
}

function normalizeDialogAnswers(answers) {
  var profile = answers.samplingProfile;
  var outputMode = answers.outputMode;
  if (profile < 0 || profile >= SAMPLING_PROFILES.length) {
    profile = 0;
  }
  if (outputMode !== OUTPUT_PITCH_CONTROL_CURVE) {
    outputMode = OUTPUT_PITCH_DELTA;
  }
  return {
    scopeMode: answers.scopeMode,
    outputMode: outputMode,
    samplingProfile: profile,
    paddingMode: answers.paddingMode,
    pureAiBaseline: answers.pureAiBaseline !== false
  };
}

function getOutputModeLabel(outputMode) {
  if (outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    return "音高控制曲线";
  }
  return "音高偏差";
}

function getErrorMessage(err) {
  if (!err) {
    return "未知错误";
  }
  return String(err.message || err);
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
  refreshStateHandles(state);
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    if (state.pureAiBaseline) {
      backupManagedPitchControls(state);
    }
    backupNotePitchModes(state);
  } else {
    backupPitchDelta(state);
  }
  state.project.newUndoRecord();

  if (state.pureAiBaseline) {
    clearExistingOutputInRanges(state);
  }
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    setNotesPitchAutoModeInRanges(state, false);
    refreshStateHandles(state);
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

function backupManagedPitchControls(state) {
  state.backupPitchControls = [];
  for (var i = 0; i < state.group.getNumPitchControls(); i++) {
    var pitchControl = state.group.getPitchControl(i);
    if (!isManagedPitchControl(pitchControl)) {
      continue;
    }
    if (!doesPitchControlOverlapRanges(pitchControl, state.ranges)) {
      continue;
    }
    state.backupPitchControls.push(pitchControl.clone());
  }
}

function clearExistingOutputInRanges(state) {
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    removeManagedPitchControlsInRanges(state.group, state.ranges);
  } else {
    clearPitchDeltaInRanges(state);
  }
}

function backupNotePitchModes(state) {
  state.backupNotePitchModes = [];
  for (var i = 0; i < state.allNotes.length; i++) {
    var note = state.allNotes[i];
    if (!doesNoteOverlapAnyRange(note, state.ranges)) {
      continue;
    }
    if (typeof note.getPitchAutoMode !== "function") {
      continue;
    }
    state.backupNotePitchModes.push({
      noteKey: getNoteKey(note),
      isAuto: note.getPitchAutoMode()
    });
  }
}

function setNotesPitchAutoModeInRanges(state, isAuto) {
  for (var i = 0; i < state.allNotes.length; i++) {
    var note = state.allNotes[i];
    if (!doesNoteOverlapAnyRange(note, state.ranges)) {
      continue;
    }
    if (typeof note.setPitchAutoMode === "function") {
      note.setPitchAutoMode(isAuto);
    }
  }
}

function restoreNotePitchModes(state) {
  refreshStateNotes(state);
  var modeMap = {};
  for (var i = 0; i < state.backupNotePitchModes.length; i++) {
    modeMap[state.backupNotePitchModes[i].noteKey] = state.backupNotePitchModes[i].isAuto;
  }
  for (var j = 0; j < state.allNotes.length; j++) {
    var note = state.allNotes[j];
    var key = getNoteKey(note);
    if (modeMap.hasOwnProperty(key) && typeof note.setPitchAutoMode === "function") {
      note.setPitchAutoMode(modeMap[key]);
    }
  }
}

function refreshStateNotes(state) {
  state.allNotes = getAllNotes(state.group);
}

function reacquireGroupReference(state) {
  if (!state || !state.project) {
    throw new Error("工程句柄不可用。");
  }
  if (state.groupRefTrackIndex < 0 || state.groupRefIndex < 0) {
    throw new Error("目标组引用位置不可用。");
  }
  var track = state.project.getTrack(state.groupRefTrackIndex);
  if (!track) {
    throw new Error("目标轨道不可用。");
  }
  var scope = track.getGroupReference(state.groupRefIndex);
  if (!scope) {
    throw new Error("目标组引用不可用。");
  }
  return scope;
}

function refreshStateHandles(state) {
  state.editor = SV.getMainEditor();
  state.scope = reacquireGroupReference(state);
  state.groupRef = state.scope;
  state.groupRefOffset = state.scope.getTimeOffset();
  state.groupPitchOffset = state.scope.getPitchOffset();
  state.group = state.scope.getTarget();
  state.pitchDelta = state.group.getParameter("pitchDelta");
  refreshStateNotes(state);
}

function clonePoints(points) {
  var cloned = [];
  for (var i = 0; i < points.length; i++) {
    cloned.push([points[i][0], points[i][1]]);
  }
  return cloned;
}

function processRangeSequentially(state, rangeIndex) {
  if (state.isFinalized) {
    return;
  }
  if (rangeIndex >= state.ranges.length) {
    finalizeSuccess(state);
    return;
  }
  try {
    refreshStateHandles(state);
  } catch (refreshErr) {
    rollbackAndFail(state, getErrorMessage(refreshErr));
    return;
  }

  var range = state.ranges[rangeIndex];
  sampleRangeAndWriteOutput(state, range, function onRangeOutputReady(err, result) {
    if (state.isFinalized) {
      return;
    }
    try {
      if (err) {
        rollbackAndFail(state, err.message || String(err));
        return;
      }

      state.writtenPoints += result.pointsWritten;
      state.outputObjectsWritten += result.outputObjectsWritten || 0;
      state.voicedFrames += result.voicedFrames;
      state.clampedPoints += result.clampedPoints;
      if (state.outputMode === OUTPUT_PITCH_DELTA && state.profile.simplifyThreshold > 0) {
        state.pitchDelta.simplify(range[0], range[1], state.profile.simplifyThreshold);
      }
      if (state.outputMode === OUTPUT_PITCH_DELTA && result.transitionGuards && result.transitionGuards.length > 0) {
        state.writtenPoints += applyConnectedTransitionGuards(state.pitchDelta, result.transitionGuards);
      }
      processRangeSequentially(state, rangeIndex + 1);
    } catch (callbackErr) {
      if (isInvalidStateError(callbackErr)) {
        rollbackAndFail(state, "写入当前区间后状态失效：" + getErrorMessage(callbackErr));
        return;
      }
      rollbackAndFail(state, getErrorMessage(callbackErr));
    }
  });
}

function sampleRangeAndWriteOutput(state, range, done) {
  done = createOnceCallback(done);
  var absoluteStart = range[0] + state.groupRefOffset;
  var absoluteEnd = range[1] + state.groupRefOffset;
  var sampleIntervalBlick = getSamplingIntervalBlick(state.timeAxis, absoluteStart, state.profile.sampleSeconds);
  var frameCount = Math.floor((absoluteEnd - absoluteStart) / sampleIntervalBlick) + 1;
  if (frameCount <= 0) {
    done(null, createEmptyWriteResult(state.outputMode));
    return;
  }

  pollComputedPitch(
    state,
    state.groupRef,
    absoluteStart,
    sampleIntervalBlick,
    frameCount,
    POLL_INTERVAL_MS,
    MAX_WAIT_MS,
    function onComputedPitchPolled(err, sampledPitch) {
      try {
        if (state.isFinalized) {
          return;
        }
        if (err) {
          if (isTimeoutError(err)) {
            done(null, writeFallbackOutputForUnavailablePitch(state, range, absoluteStart, sampleIntervalBlick));
            return;
          }
          done(err);
          return;
        }
        var writeResult;
        try {
          writeResult = writeOutputWithRetry(state, range, absoluteStart, sampleIntervalBlick, sampledPitch);
        } catch (writeErr) {
          done(writeErr);
          return;
        }
        if (writeResult.voicedFrames <= 0) {
          done(null, writeFallbackOutputForUnavailablePitch(state, range, absoluteStart, sampleIntervalBlick));
          return;
        }
        done(null, writeResult);
      } catch (callbackErr) {
        try {
          done(callbackErr);
        } catch (doneErr) {
          handleFatalScriptError(doneErr);
        }
      }
    }
  );
}

function createOnceCallback(fn) {
  var called = false;
  return function() {
    if (called) {
      return;
    }
    called = true;
    return fn.apply(null, arguments);
  };
}

function writeOutputWithRetry(state, range, absoluteStart, sampleIntervalBlick, sampledPitch) {
  try {
    return writeOutputOnce(state, range, absoluteStart, sampleIntervalBlick, sampledPitch);
  } catch (err) {
    if (!isInvalidStateError(err)) {
      throw err;
    }
    refreshStateHandles(state);
    return writeOutputOnce(state, range, absoluteStart, sampleIntervalBlick, sampledPitch);
  }
}

function writeOutputOnce(state, range, absoluteStart, sampleIntervalBlick, sampledPitch) {
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    return writePitchControlCurveFromSampledPitch(state, range, absoluteStart, sampleIntervalBlick, sampledPitch);
  }
  return writePitchDeltaFromSampledPitch(state, range, absoluteStart, sampleIntervalBlick, sampledPitch);
}

function createEmptyWriteResult(outputMode) {
  return {
    pointsWritten: 0,
    outputObjectsWritten: 0,
    voicedFrames: 0,
    clampedPoints: 0,
    transitionGuards: outputMode === OUTPUT_PITCH_DELTA ? [] : undefined
  };
}

function writeFallbackOutputForUnavailablePitch(state, range, absoluteStart, sampleIntervalBlick) {
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    return createEmptyWriteResult(state.outputMode);
  }
  // If rendered pitch is unavailable (e.g. trial voicebank limits or deleted lyrics),
  // fill this range with zero anchors instead of failing the whole script.
  return writePitchDeltaFromSampledPitch(state, range, absoluteStart, sampleIntervalBlick, []);
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

function pollComputedPitch(state, groupRef, absStart, intervalBlick, frameCount, pollMs, maxWaitMs, callback) {
  var startedMs = new Date().getTime();
  var settled = false;

  function finishOnce(err, pitches) {
    if (settled) {
      return;
    }
    settled = true;
    try {
      callback(err, pitches);
    } catch (callbackErr) {
      handleFatalScriptError(callbackErr);
    }
  }

  function tryFetch() {
    if (settled || state.isFinalized) {
      return;
    }
    try {
      var pitches = SV.getComputedPitchForGroup(
        groupRef,
        absStart,
        intervalBlick,
        frameCount
      );
      if (pitches && pitches.length > 0) {
        finishOnce(null, pitches);
        return;
      }

      var elapsed = new Date().getTime() - startedMs;
      if (elapsed >= maxWaitMs) {
        finishOnce(new Error("等待计算音高超时。"));
        return;
      }

      SV.setTimeout(pollMs, function onComputedPitchPollTimer() {
        try {
          if (settled || state.isFinalized) {
            return;
          }
          tryFetch();
        } catch (timerErr) {
          finishOnce(timerErr);
        }
      });
    } catch (err) {
      finishOnce(err);
    }
  }

  tryFetch();
}

function isTimeoutError(err) {
  if (!err) {
    return false;
  }
  return String(err.message || err) === "等待计算音高超时。";
}

function isInvalidStateError(err) {
  if (!err) {
    return false;
  }
  return String(err.message || err).toLowerCase().indexOf("invalid state") >= 0;
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

  // Only anchor note edges that touch a rest or the extraction boundary.
  // For connected note junctions, forcing head/tail anchors creates large
  // discontinuities that can overshoot under pitchDelta interpolation.
  for (var n = 0; n < rangeNotes.length; n++) {
    var rangeNote = rangeNotes[n];
    var noteStart = Math.max(range[0], rangeNote.getOnset());
    var noteEndExclusive = Math.min(range[1], rangeNote.getEnd());
    if (noteEndExclusive <= noteStart) {
      continue;
    }
    var edge = noteEdgeMap[getNoteKey(rangeNote)];
    var prevNote = n > 0 ? rangeNotes[n - 1] : null;
    var nextNote = n + 1 < rangeNotes.length ? rangeNotes[n + 1] : null;
    var prevEnd = prevNote ? Math.min(range[1], prevNote.getEnd()) : range[0];
    var nextStart = nextNote ? Math.max(range[0], nextNote.getOnset()) : range[1];
    var shouldAnchorHead = noteStart <= range[0] || prevEnd < noteStart;
    var shouldAnchorTail = noteEndExclusive >= range[1] || nextStart > noteEndExclusive;
    var noteEndInclusive = noteEndExclusive - 1;

    if (shouldAnchorHead) {
      addPoint(noteStart, edge ? edge.firstDelta : 0);
    }
    if (shouldAnchorTail) {
      addPoint(noteEndInclusive, edge ? edge.lastDelta : 0);
    }
  }

  // Force pitchDelta to 0 in non-note regions so disconnected notes do not bridge.
  addRestZeroAnchors(range, rangeNotes, addPoint);

  return {
    pointsWritten: pointsWritten,
    outputObjectsWritten: 0,
    voicedFrames: voicedFrames,
    clampedPoints: clampedPoints,
    transitionGuards: buildConnectedTransitionGuards(range, rangeNotes, noteEdgeMap)
  };
}

function writePitchControlCurveFromSampledPitch(state, range, absoluteStart, sampleIntervalBlick, sampledPitch) {
  var voicedFrames = 0;
  var noteEdgeMap = {};
  var rangeNotes = getNotesIntersectingRange(state.allNotes, range);
  var absolutePoints = [];
  var noteIndex = findNoteIndexAtOrBefore(state.allNotes, range[0]);
  if (noteIndex < 0) {
    noteIndex = 0;
  }

  for (var i = 0; i < sampledPitch.length; i++) {
    var rawPitch = sampledPitch[i];
    if (rawPitch === null || rawPitch === undefined || !isFinite(rawPitch) || rawPitch <= 0) {
      continue;
    }

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
    var groupRelativePitch = rawPitch - state.groupPitchOffset;
    absolutePoints.push([localBlick, groupRelativePitch]);

    var noteKey = getNoteKey(note);
    if (!noteEdgeMap[noteKey]) {
      noteEdgeMap[noteKey] = {
        firstBlick: localBlick,
        firstPitch: groupRelativePitch,
        lastBlick: localBlick,
        lastPitch: groupRelativePitch
      };
    } else {
      if (localBlick < noteEdgeMap[noteKey].firstBlick) {
        noteEdgeMap[noteKey].firstBlick = localBlick;
        noteEdgeMap[noteKey].firstPitch = groupRelativePitch;
      }
      if (localBlick > noteEdgeMap[noteKey].lastBlick) {
        noteEdgeMap[noteKey].lastBlick = localBlick;
        noteEdgeMap[noteKey].lastPitch = groupRelativePitch;
      }
    }
  }

  if (absolutePoints.length === 0) {
    return createEmptyWriteResult(state.outputMode);
  }

  var anchoredPoints = addPitchCurveBoundaryAnchors(range, rangeNotes, noteEdgeMap, absolutePoints);
  var dedupedPoints = dedupePitchCurvePoints(anchoredPoints);
  if (dedupedPoints.length === 0) {
    return createEmptyWriteResult(state.outputMode);
  }

  var curve = buildPitchControlCurveFromAbsolutePoints(state, range, dedupedPoints);
  state.group.addPitchControl(curve);

  return {
    pointsWritten: dedupedPoints.length,
    outputObjectsWritten: 1,
    voicedFrames: voicedFrames,
    clampedPoints: 0
  };
}

function addPitchCurveBoundaryAnchors(range, rangeNotes, noteEdgeMap, points) {
  var anchored = points.slice();
  for (var i = 0; i < rangeNotes.length; i++) {
    var note = rangeNotes[i];
    var edge = noteEdgeMap[getNoteKey(note)];
    if (!edge) {
      continue;
    }
    var noteStart = Math.max(range[0], note.getOnset());
    var noteEndExclusive = Math.min(range[1], note.getEnd());
    if (noteEndExclusive <= noteStart) {
      continue;
    }
    anchored.push([noteStart, edge.firstPitch]);
    anchored.push([noteEndExclusive - 1, edge.lastPitch]);
  }
  return anchored;
}

function dedupePitchCurvePoints(points) {
  if (points.length === 0) {
    return [];
  }

  points.sort(function(a, b) {
    if (a[0] !== b[0]) {
      return a[0] - b[0];
    }
    return a[1] - b[1];
  });

  var deduped = [];
  for (var i = 0; i < points.length; i++) {
    var blick = Math.round(points[i][0]);
    var value = points[i][1];
    if (deduped.length > 0 && deduped[deduped.length - 1][0] === blick) {
      deduped[deduped.length - 1][1] = value;
    } else {
      deduped.push([blick, value]);
    }
  }
  return deduped;
}

function buildPitchControlCurveFromAbsolutePoints(state, range, points) {
  var bounds = getPitchCurveBounds(points);
  var anchorPosition = Math.round((bounds.minBlick + bounds.maxBlick) / 2);
  var anchorPitch = (bounds.minPitch + bounds.maxPitch) / 2.0;
  var relativePoints = [];

  for (var i = 0; i < points.length; i++) {
    relativePoints.push([
      points[i][0] - anchorPosition,
      points[i][1] - anchorPitch
    ]);
  }

  var curve = SV.create("PitchControlCurve");
  curve.setPosition(anchorPosition);
  curve.setPitch(anchorPitch);
  curve.setPoints(relativePoints);
  curve.setScriptData(MANAGED_PITCH_CONTROL_KEY, {
    outputMode: "PitchControlCurve",
    sessionId: state.sessionId,
    range: [range[0], range[1]]
  });
  return curve;
}

function getPitchCurveBounds(points) {
  var minBlick = points[0][0];
  var maxBlick = points[0][0];
  var minPitch = points[0][1];
  var maxPitch = points[0][1];

  for (var i = 1; i < points.length; i++) {
    if (points[i][0] < minBlick) {
      minBlick = points[i][0];
    }
    if (points[i][0] > maxBlick) {
      maxBlick = points[i][0];
    }
    if (points[i][1] < minPitch) {
      minPitch = points[i][1];
    }
    if (points[i][1] > maxPitch) {
      maxPitch = points[i][1];
    }
  }

  return {
    minBlick: minBlick,
    maxBlick: maxBlick,
    minPitch: minPitch,
    maxPitch: maxPitch
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

function doesNoteOverlapAnyRange(note, ranges) {
  var noteRange = [note.getOnset(), note.getEnd()];
  for (var i = 0; i < ranges.length; i++) {
    if (rangesOverlap(noteRange, ranges[i])) {
      return true;
    }
  }
  return false;
}

function isManagedPitchControl(pitchControl) {
  return !!(pitchControl
    && typeof pitchControl.hasScriptData === "function"
    && pitchControl.hasScriptData(MANAGED_PITCH_CONTROL_KEY));
}

function getManagedPitchControlInfo(pitchControl) {
  if (!isManagedPitchControl(pitchControl)) {
    return null;
  }
  return pitchControl.getScriptData(MANAGED_PITCH_CONTROL_KEY);
}

function doesPitchControlOverlapRanges(pitchControl, ranges) {
  var controlRange = getPitchControlRange(pitchControl);
  for (var i = 0; i < ranges.length; i++) {
    if (rangesOverlap(controlRange, ranges[i])) {
      return true;
    }
  }
  return false;
}

function getPitchControlRange(pitchControl) {
  var position = Math.round(pitchControl.getPosition());
  if (typeof pitchControl.getPoints !== "function") {
    return [position, position + 1];
  }

  var points = pitchControl.getPoints();
  if (!points || points.length === 0) {
    return [position, position + 1];
  }

  var minTime = position + points[0][0];
  var maxTime = minTime;
  for (var i = 1; i < points.length; i++) {
    var absoluteTime = position + points[i][0];
    if (absoluteTime < minTime) {
      minTime = absoluteTime;
    }
    if (absoluteTime > maxTime) {
      maxTime = absoluteTime;
    }
  }
  return [Math.floor(minTime), Math.ceil(maxTime) + 1];
}

function removeManagedPitchControlsInRanges(group, ranges, sessionId) {
  for (var i = group.getNumPitchControls() - 1; i >= 0; i--) {
    var pitchControl = group.getPitchControl(i);
    var info = getManagedPitchControlInfo(pitchControl);
    if (!info) {
      continue;
    }
    if (sessionId && info.sessionId !== sessionId) {
      continue;
    }
    if (!doesPitchControlOverlapRanges(pitchControl, ranges)) {
      continue;
    }
    group.removePitchControl(i);
  }
}

function buildConnectedTransitionGuards(range, rangeNotes, noteEdgeMap) {
  var guards = [];
  for (var i = 0; i + 1 < rangeNotes.length; i++) {
    var prevNote = rangeNotes[i];
    var nextNote = rangeNotes[i + 1];
    var prevEnd = Math.min(range[1], prevNote.getEnd());
    var nextStart = Math.max(range[0], nextNote.getOnset());
    if (nextStart > prevEnd) {
      continue;
    }

    var prevEdge = noteEdgeMap[getNoteKey(prevNote)];
    var nextEdge = noteEdgeMap[getNoteKey(nextNote)];
    if (!prevEdge || !nextEdge) {
      continue;
    }

    guards.push({
      boundary: Math.round(prevEnd),
      prevStart: Math.max(range[0], prevNote.getOnset()),
      prevEnd: prevEnd,
      nextStart: nextStart,
      nextEnd: Math.min(range[1], nextNote.getEnd()),
      prevDelta: prevEdge.lastDelta,
      nextDelta: nextEdge.firstDelta
    });
  }
  return guards;
}

function applyConnectedTransitionGuards(automation, guards) {
  var added = 0;
  for (var i = 0; i < guards.length; i++) {
    added += addTransitionGuard(automation, guards[i]);
  }
  return added;
}

function addTransitionGuard(automation, guard) {
  var added = 0;
  var prevPositions = [guard.boundary - 2, guard.boundary - 1];
  var nextPositions = [guard.boundary, guard.boundary + 1];

  for (var i = 0; i < prevPositions.length; i++) {
    if (isBlickInHalfOpenRange(prevPositions[i], guard.prevStart, guard.prevEnd)) {
      if (automation.add(prevPositions[i], guard.prevDelta)) {
        added++;
      }
    }
  }
  for (var j = 0; j < nextPositions.length; j++) {
    if (isBlickInHalfOpenRange(nextPositions[j], guard.nextStart, guard.nextEnd)) {
      if (automation.add(nextPositions[j], guard.nextDelta)) {
        added++;
      }
    }
  }
  return added;
}

function isBlickInHalfOpenRange(blick, startInclusive, endExclusive) {
  return startInclusive <= blick && blick < endExclusive;
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
  if (state.isFinalized) {
    return;
  }
  state.isFinalized = true;
  var rollbackError = null;
  try {
    refreshStateHandles(state);
    if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
      removeManagedPitchControlsInRanges(state.group, state.ranges, state.sessionId);
      restoreNotePitchModes(state);
      for (var i = 0; i < state.backupPitchControls.length; i++) {
        state.group.addPitchControl(state.backupPitchControls[i]);
      }
    } else {
      for (var j = 0; j < state.ranges.length; j++) {
        state.pitchDelta.remove(state.ranges[j][0], state.ranges[j][1]);
      }
      for (var m = 0; m < state.backupPointsByRange.length; m++) {
        var backup = state.backupPointsByRange[m];
        for (var k = 0; k < backup.points.length; k++) {
          state.pitchDelta.add(backup.points[k][0], backup.points[k][1]);
        }
      }
    }
  } catch (err) {
    rollbackError = err;
  }

  finishThenShowMessage(
    SCRIPT_TITLE,
    formatFailureMessage(reason, rollbackError)
  );
}

function finalizeSuccess(state) {
  if (state.isFinalized) {
    return;
  }
  state.isFinalized = true;
  var elapsed = (new Date().getTime() - state.startTimeMs) / 1000.0;
  if (state.outputMode === OUTPUT_PITCH_CONTROL_CURVE) {
    finishThenShowMessage(
      SCRIPT_TITLE,
      formatMessage(
        "提取完成。\n输出模式：%s\n区间数：%d\n采样帧数：%d\n新增曲线：%d\n曲线点数：%d\n切换为手动音高的音符数：%d\n耗时：%.2f 秒",
        [getOutputModeLabel(state.outputMode), state.ranges.length, state.voicedFrames, state.outputObjectsWritten, state.writtenPoints, state.backupNotePitchModes.length, elapsed]
      )
    );
  } else {
    var visiblePoints = state.writtenPoints;
    try {
      visiblePoints = countPointsInRanges(state.pitchDelta, state.ranges);
    } catch (countErr) {}
    finishThenShowMessage(
      SCRIPT_TITLE,
      formatMessage(
        "提取完成。\n输出模式：%s\n区间数：%d\n采样帧数：%d\n可见控制点：%d\n限幅帧数：%d\n耗时：%.2f 秒",
        [getOutputModeLabel(state.outputMode), state.ranges.length, state.voicedFrames, visiblePoints, state.clampedPoints, elapsed]
      )
    );
  }
}

function formatFailureMessage(reason, rollbackError) {
  var message = formatMessage("提取失败，已回滚。\n原因：%s", [reason]);
  if (rollbackError) {
    message += "\n回滚错误：" + getErrorMessage(rollbackError);
  }
  return message;
}

function handleFatalScriptError(err) {
  finishThenShowMessage(SCRIPT_TITLE, "脚本执行失败：\n" + getErrorMessage(err));
}

function finishThenShowMessage(title, message) {
  var finishError = null;
  try {
    SV.finish();
  } catch (err) {
    finishError = err;
  }
  if (finishError) {
    message += "\n\n结束脚本时发生错误：" + getErrorMessage(finishError);
  }
  SV.showMessageBox(title, message);
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
