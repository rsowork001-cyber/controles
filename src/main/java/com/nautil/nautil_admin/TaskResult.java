package com.nautil;

import java.util.ArrayList;
import java.util.List;

public class TaskResult {

    private boolean success;
    private String message;
    private List<String> logs = new ArrayList<>();

    public TaskResult() {}

    public TaskResult(boolean success, String message) {
        this.success = success;
        this.message = message;
    }

    public void addLog(String log) {
        this.logs.add(log);
    }

    public boolean isSuccess() { return success; }
    public void setSuccess(boolean success) { this.success = success; }

    public String getMessage() { return message; }
    public void setMessage(String message) { this.message = message; }

    public List<String> getLogs() { return logs; }
    public void setLogs(List<String> logs) { this.logs = logs; }
}
