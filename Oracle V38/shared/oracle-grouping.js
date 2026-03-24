// oracle-grouping.js — Task grouping logic (Drive files, Slack channels, Tags)
// Used by both newtab.js and sidepanel.js

(function () {
  'use strict';

  const { isValidTag } = window.Oracle;
  const { isSlackLink, extractDriveFileId, getCleanDriveFileUrl } = window.OracleIcons;

  // Extract a reasonable file name from a task
  function extractFileNameFromTask(task) {
    if (task.task_title) return task.task_title;
    const name = task.task_name || '';
    const quotedMatch = name.match(/'([^']+)'|"([^"]+)"/);
    if (quotedMatch) return quotedMatch[1] || quotedMatch[2];
    return name.substring(0, 50) + (name.length > 50 ? '...' : '');
  }

  // Group tasks by Google Drive file ID
  function groupTasksByDriveFile(tasks) {
    const driveGroups = {};
    const nonDriveTasks = [];

    tasks.forEach(task => {
      const fileId = extractDriveFileId(task.message_link);
      if (fileId) {
        if (!driveGroups[fileId]) {
          const cleanFileUrl = getCleanDriveFileUrl(task.message_link, fileId);
          driveGroups[fileId] = {
            fileId,
            fileUrl: cleanFileUrl,
            tasks: [],
            groupTitle: task.participant_text || task.task_title || extractFileNameFromTask(task),
            latestUpdate: task.updated_at || task.created_at
          };
        }
        driveGroups[fileId].tasks.push(task);
        const taskTime = new Date(task.updated_at || task.created_at);
        const groupTime = new Date(driveGroups[fileId].latestUpdate);
        if (taskTime > groupTime) {
          driveGroups[fileId].latestUpdate = task.updated_at || task.created_at;
        }
      } else {
        nonDriveTasks.push(task);
      }
    });

    return { driveGroups, nonDriveTasks };
  }

  // Group tasks by Slack channel (using participant_text)
  function groupTasksBySlackChannel(tasks) {
    const slackGroups = {};
    const nonSlackTasks = [];

    tasks.forEach(task => {
      const isDM = task.participant_text && task.participant_text.toLowerCase().startsWith('dm with');
      if (isSlackLink(task.message_link) && task.participant_text && !isDM) {
        const channelKey = task.participant_text.trim();
        if (!slackGroups[channelKey]) {
          slackGroups[channelKey] = {
            channelName: channelKey,
            tasks: [],
            latestUpdate: task.updated_at || task.created_at
          };
        }
        slackGroups[channelKey].tasks.push(task);
        const taskTime = new Date(task.updated_at || task.created_at);
        const groupTime = new Date(slackGroups[channelKey].latestUpdate);
        if (taskTime > groupTime) {
          slackGroups[channelKey].latestUpdate = task.updated_at || task.created_at;
        }
      } else {
        nonSlackTasks.push(task);
      }
    });

    // Only keep groups with >1 task; singletons go back to non-grouped
    const filteredSlackGroups = {};
    Object.entries(slackGroups).forEach(([key, group]) => {
      if (group.tasks.length > 1) {
        filteredSlackGroups[key] = group;
      } else {
        nonSlackTasks.push(...group.tasks);
      }
    });

    return { slackGroups: filteredSlackGroups, nonSlackTasks };
  }

  // Group tasks by tag (primary tag only — first valid tag)
  function groupTasksByTag(tasks) {
    const tagGroups = {};
    const untaggedTasks = [];

    tasks.forEach(task => {
      const taskTags = (task.tags || []).filter(isValidTag);
      if (taskTags.length === 0) {
        untaggedTasks.push(task);
        return;
      }
      const primaryTag = taskTags[0].trim();
      if (!isValidTag(primaryTag)) {
        untaggedTasks.push(task);
        return;
      }
      if (!tagGroups[primaryTag]) {
        tagGroups[primaryTag] = {
          tagName: primaryTag,
          tasks: [],
          latestUpdate: task.updated_at || task.created_at
        };
      }
      tagGroups[primaryTag].tasks.push(task);
      const taskTime = new Date(task.updated_at || task.created_at);
      const groupTime = new Date(tagGroups[primaryTag].latestUpdate);
      if (taskTime > groupTime) {
        tagGroups[primaryTag].latestUpdate = task.updated_at || task.created_at;
      }
    });

    return { tagGroups, untaggedTasks };
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleGrouping = {
    extractFileNameFromTask,
    groupTasksByDriveFile,
    groupTasksBySlackChannel,
    groupTasksByTag,
  };

})();
