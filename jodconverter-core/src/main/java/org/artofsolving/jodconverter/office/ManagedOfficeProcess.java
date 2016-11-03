//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//	-> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//	-> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.office;

import com.sun.star.frame.XDesktop;
import com.sun.star.lang.DisposedException;

import java.net.ConnectException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.logging.Level;
import java.util.logging.Logger;

class ManagedOfficeProcess {

	private static final Integer EXIT_CODE_NEW_INSTALLATION = 81;

	private final ManagedOfficeProcessSettings settings;

	private final OfficeProcess process;
	private final OfficeConnection connection;

	private ExecutorService executor = Executors.newSingleThreadExecutor(new NamedThreadFactory("OfficeProcessThread"));

	private final Logger logger = Logger.getLogger(getClass().getName());

	public ManagedOfficeProcess(ManagedOfficeProcessSettings settings) throws OfficeException {
		this.settings = settings;
		process = new OfficeProcess(settings.getOfficeHome(),
				settings.getUnoUrl(),
				settings.getRunAsArgs(),
				settings.getTemplateProfileDir(),
				settings.getWorkDir(),
				settings.getProcessManager(),
				settings.isKillExistingProcess(),
				settings.getStartupWatcherTimeout());
		connection = new OfficeConnection(settings.getUnoUrl());
	}

	public OfficeConnection getConnection() {
		return connection;
	}

	public void startAndWait() throws OfficeException {
		Future<?> future = executor.submit(new Runnable() {
			public void run() {
				doStartProcessAndConnect();
			}
		});
		try {
			future.get();
		} catch (Exception exception) {
			throw new OfficeException("failed to start and connect", exception);
		}
	}

	public void stopAndWait() throws OfficeException {
		Future<?> future = executor.submit(new Runnable() {
			public void run() {
				doStopProcess();
			}
		});
		try {
			future.get();
		} catch (Exception exception) {
			throw new OfficeException("failed to stop", exception);
		}
	}

	public void restartAndWait() {
		Future<?> future = executor.submit(new Runnable() {
			public void run() {
				doStopProcess();
				doStartProcessAndConnect();
			}
		});
		try {
			future.get();
		} catch (Exception exception) {
			throw new OfficeException("failed to restart", exception);
		}
	}

	public void restartDueToTaskTimeout() {
		executor.execute(new Runnable() {
			public void run() {
				doTerminateProcess();
				// will cause unexpected disconnection and subsequent restart
			}
		});
	}

	public void restartDueToLostConnection() {
		executor.execute(new Runnable() {
			public void run() {
				try {
					doEnsureProcessExited();
					doStartProcessAndConnect();
				} catch (OfficeException officeException) {
					logger.log(Level.SEVERE, "could not restart process", officeException);
				}
			}
		});
	}

	private void doStartProcessAndConnect() throws OfficeException {
		try {
			process.start();
			new Retryable() {
				protected void attempt() throws TemporaryException, Exception {
					try {
						connection.connect();
					} catch (ConnectException connectException) {
							throw new TemporaryException(connectException);
					}
				}
			}.execute(settings.getRetryInterval(), settings.getRetryTimeout());
		} catch (Exception exception) {
			throw new OfficeException("could not establish connection", exception);
		}
	}

	private void doStopProcess() {
		try {
			XDesktop desktop = OfficeUtils.cast(XDesktop.class, connection.getService(OfficeUtils.SERVICE_DESKTOP));
			desktop.terminate();
		} catch (DisposedException disposedException) {
			// expected
		} catch (Exception exception) {
			// in case we can't get hold of the desktop
			doTerminateProcess();
		}
		doEnsureProcessExited();
	}

	private void doEnsureProcessExited() throws OfficeException {
		try {
			int exitCode = process.getExitCode(settings.getRetryInterval(), settings.getRetryTimeout());
			logger.info("process exited with code " + exitCode);
		} catch (RetryTimeoutException retryTimeoutException) {
			doTerminateProcess();
		}
		process.deleteProfileDir();
	}

	private void doTerminateProcess() throws OfficeException {
		try {
			int exitCode = process.forciblyTerminate(settings.getRetryInterval(), settings.getRetryTimeout());
			logger.info("process forcibly terminated with code " + exitCode);
		} catch (Exception exception) {
			throw new OfficeException("could not terminate process", exception);
		}
	}

    boolean isConnected() {
        return connection.isConnected();
    }


	@Override
	public String toString() {
		final StringBuilder sb = new StringBuilder("ManagedOfficeProcess{");
		sb.append("settings=").append(settings);
		sb.append(", process=").append(process);
		sb.append('}');
		return sb.toString();
	}
}
