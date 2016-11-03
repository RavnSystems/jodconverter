//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.office;

import org.apache.commons.io.FileUtils;
import org.artofsolving.jodconverter.process.ProcessManager;
import org.artofsolving.jodconverter.process.ProcessQuery;
import org.artofsolving.jodconverter.util.PlatformUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import static org.artofsolving.jodconverter.process.ProcessManager.PID_NOT_FOUND;
import static org.artofsolving.jodconverter.process.ProcessManager.PID_UNKNOWN;

class OfficeProcess {

    private final File officeHome;
    private final UnoUrl unoUrl;
    private final String[] runAsArgs;
    private final File templateProfileDir;
    private final File instanceProfileDir;
    private final ProcessManager processManager;
    private final boolean killExistingProcess;
    private final int startupWatcherTimeout;

    private Process process;
    private long pid = PID_UNKNOWN;

    private final Logger logger = Logger.getLogger(getClass().getName());

    public OfficeProcess(File officeHome, UnoUrl unoUrl, String[] runAsArgs, File templateProfileDir, File workDir, ProcessManager processManager, boolean killExistingProcess, int startupWatcherTimeout) {
        this.officeHome = officeHome;
        this.unoUrl = unoUrl;
        this.runAsArgs = runAsArgs;
        this.templateProfileDir = templateProfileDir;
        this.killExistingProcess = killExistingProcess;
        this.startupWatcherTimeout = startupWatcherTimeout;
        this.instanceProfileDir = getInstanceProfileDir(workDir, unoUrl);
        this.processManager = processManager;
    }

    public void start() throws IOException {
        start(false);
    }

    private void start(boolean restart) throws IOException {
        ProcessQuery processQuery = new ProcessQuery("soffice.bin", unoUrl.getAcceptString());
        long existingPid = processManager.findPid(processQuery);
    	if (!(existingPid == PID_NOT_FOUND || existingPid == PID_UNKNOWN) && killExistingProcess) {
            logger.warning(String.format("a process with acceptString '%s' is already running; pid %s",
                    unoUrl.getAcceptString(), existingPid));
            processManager.kill(null, existingPid);
            waitForProcessToDie(existingPid);
        }
        if(!(existingPid == PID_NOT_FOUND || existingPid == PID_UNKNOWN)){
			throw new IllegalStateException(String.format("a process with acceptString '%s' is already running; pid %d",
			        unoUrl.getAcceptString(), existingPid));
        }
    	if (!restart) {
    	    prepareInstanceProfileDir();
    	}
        List<String> command = new ArrayList<String>();
        File executable = OfficeUtils.getOfficeExecutable(officeHome);
        if (runAsArgs != null) {
        	command.addAll(Arrays.asList(runAsArgs));
        }
        command.add(executable.getAbsolutePath());
        command.add("-accept=" + unoUrl.getAcceptString() + ";urp;");
        command.add("-env:UserInstallation=" + OfficeUtils.toUrl(instanceProfileDir));
        command.add("-headless");
        command.add("-nocrashreport");
        command.add("-nodefault");
        command.add("-nofirststartwizard");
        command.add("-nolockcheck");
        command.add("-nologo");
        command.add("-norestore");
        ProcessBuilder processBuilder = new ProcessBuilder(command);
        if (PlatformUtils.isWindows()) {
            addBasisAndUrePaths(processBuilder);
        }
        logger.info(String.format("starting process with acceptString '%s' and profileDir '%s'", unoUrl, instanceProfileDir));
        process = processBuilder.start();
        logger.info("started process" + (pid != PID_UNKNOWN ? "; pid = " + pid : ""));
        if(isNewInstallation()) {
            waitForProcessToDie();
            logger.warning("Restarting OOo after code 81 ...");
            process = processBuilder.start();
        }
        pid = getPid(processQuery);
        manageInputStreams(process);
    }


    private boolean isNewInstallation() {
        int exitValue = 0;
        boolean exited = false;
        final int timeout = startupWatcherTimeout * 2;
        for (int i = 0; i < timeout; i++) {
            try {
                // wait for process to start
                if(i!=0)
                    Thread.sleep(500);
            } catch (Exception ignore) {
            }
            try {
                exitValue = process.exitValue();
                // process is already dead, no need to wait longer ...
                exited = true;
                break;
            } catch (IllegalThreadStateException e) {
                // process is still up
            }
        }

        if (exited && exitValue==81)
                return true;
        return false;
    }

    private void waitForProcessToDie(long pid){
        for(int i=0; i<60*4; i++){
            if(!isRunning())
                return;
            try {
                Thread.sleep(250);
            } catch (InterruptedException e) {
                break;
            }
        }
        throw new OfficeException("Unable to kill process with pid "+pid);
    }


    //wait for one minute
    private void waitForProcessToDie() {
        waitForProcessToDie(this.pid);
    }

    private void manageInputStreams(final Process process){
        final Thread stdin = new Thread(() -> logStream(process.getInputStream()));
        stdin.setDaemon(true);
        stdin.start();

        final Thread stderr = new Thread(() -> logStream(process.getErrorStream()));
        stderr.setDaemon(true);
        stderr.start();
    }

    private void logStream(InputStream inputStream) {
        try (InputStreamReader inputStreamReader = new InputStreamReader(inputStream);
             BufferedReader bufferedReader = new BufferedReader(inputStreamReader)) {
            String line;
            while((line = bufferedReader.readLine())!=null)
                logger.info("soffice: "+ line);
        } catch (IOException e) {
            logger.info(e.getMessage());
        }
    }


        /**
     * @return the process pid if started within 1 minute. Otherwise, it throws
     * @throws IOException
     */
    private long getPid(ProcessQuery processQuery) throws IOException {
        for (int i = 0; i <= 60; i++) {
            final long foundPid = processManager.findPid(processQuery);
            if (foundPid != PID_NOT_FOUND)
                return foundPid;
            if (i != 60)
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException e) {
                    break;
                }
        }
        throw new IllegalStateException(String.format("process with acceptString '%s' started but its pid could not be found",
                unoUrl.getAcceptString()));

    }

    private File getInstanceProfileDir(File workDir, UnoUrl unoUrl) {
        String dirName = ".jodconverter_" + unoUrl.getAcceptString().replace(',', '_').replace('=', '-');
        return new File(workDir, dirName);
    }

    private void prepareInstanceProfileDir() throws OfficeException {
        if (instanceProfileDir.exists()) {
            logger.warning(String.format("profile dir '%s' already exists; deleting", instanceProfileDir));
            deleteProfileDir();
        }
        if (templateProfileDir != null) {
            try {
                FileUtils.copyDirectory(templateProfileDir, instanceProfileDir);
            } catch (IOException ioException) {
                throw new OfficeException("failed to create profileDir", ioException);
            }
        }
    }

    public void deleteProfileDir() {
        if (instanceProfileDir != null) {
            try {
                FileUtils.deleteDirectory(instanceProfileDir);
            } catch (IOException ioException) {
                File oldProfileDir = new File(instanceProfileDir.getParentFile(), instanceProfileDir.getName() + ".old." + System.currentTimeMillis());
                if (instanceProfileDir.renameTo(oldProfileDir)) {
                    logger.warning("could not delete profileDir: " + ioException.getMessage() + "; renamed it to " + oldProfileDir);
                } else {
                    logger.severe("could not delete profileDir: " + ioException.getMessage());
                }
            }
        }
    }

    private void addBasisAndUrePaths(ProcessBuilder processBuilder) throws IOException {
        // see http://wiki.services.openoffice.org/wiki/ODF_Toolkit/Efforts/Three-Layer_OOo
        File basisLink = new File(officeHome, "basis-link");
        if (!basisLink.isFile()) {
            logger.fine("no %OFFICE_HOME%/basis-link found; assuming it's OOo 2.x and we don't need to append URE and Basic paths");
            return;
        }
        String basisLinkText = FileUtils.readFileToString(basisLink).trim();
        File basisHome = new File(officeHome, basisLinkText);
        File basisProgram = new File(basisHome, "program");
        File ureLink = new File(basisHome, "ure-link");
        String ureLinkText = FileUtils.readFileToString(ureLink).trim();
        File ureHome = new File(basisHome, ureLinkText);
        File ureBin = new File(ureHome, "bin");
        Map<String,String> environment = processBuilder.environment();
        // Windows environment variables are case insensitive but Java maps are not :-/
        // so let's make sure we modify the existing key
        String pathKey = "PATH";
        for (String key : environment.keySet()) {
            if ("PATH".equalsIgnoreCase(key)) {
                pathKey = key;
            }
        }
        String path = environment.get(pathKey) + ";" + ureBin.getAbsolutePath() + ";" + basisProgram.getAbsolutePath();
        logger.fine(String.format("setting %s to \"%s\"", pathKey, path));
        environment.put(pathKey, path);
    }

    public boolean isRunning() {
        if (process == null) {
            return false;
        }
        return getExitCode() == null || processManager.isRunning(pid);
    }

    private class ExitCodeRetryable extends Retryable {
        
        private int exitCode;
        
        protected void attempt() throws TemporaryException, Exception {
            try {
                exitCode = process.exitValue();
            } catch (IllegalThreadStateException illegalThreadStateException) {
                throw new TemporaryException(illegalThreadStateException);
            }
        }
        
        public int getExitCode() {
            return exitCode;
        }

    }

    public Integer getExitCode() {
        try {
            return process.exitValue();
        } catch (IllegalThreadStateException exception) {
            return null;
        }
    }

    public int getExitCode(long retryInterval, long retryTimeout) throws RetryTimeoutException {
        try {
            ExitCodeRetryable retryable = new ExitCodeRetryable();
            retryable.execute(retryInterval, retryTimeout);
            return retryable.getExitCode();
        } catch (RetryTimeoutException retryTimeoutException) {
            throw retryTimeoutException;
        } catch (Exception exception) {
            throw new OfficeException("could not get process exit code", exception);
        }
    }

    public int forciblyTerminate(long retryInterval, long retryTimeout) throws IOException, RetryTimeoutException {
        logger.info(String.format("trying to forcibly terminate process: '" + unoUrl + "'" + (pid != PID_UNKNOWN ? " (pid " + pid  + ")" : "")));
        processManager.kill(process, pid);
        return getExitCode(retryInterval, retryTimeout);
    }

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder("OfficeProcess{");
        sb.append("officeHome=").append(officeHome);
        sb.append(", unoUrl=").append(unoUrl);
        sb.append(", runAsArgs=").append(Arrays.toString(runAsArgs));
        sb.append(", templateProfileDir=").append(templateProfileDir);
        sb.append(", instanceProfileDir=").append(instanceProfileDir);
        sb.append(", processManager=").append(processManager);
        sb.append(", killExistingProcess=").append(killExistingProcess);
        sb.append(", startupWatcherTimeout=").append(startupWatcherTimeout);
        sb.append(", process=").append(process);
        sb.append(", pid=").append(pid);
        sb.append('}');
        return sb.toString();
    }
}
