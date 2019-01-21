/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.sciforma.pharma.exeception;

/**
 *
 * @author lahoudie
 */
public class TechnicalException extends RuntimeException {

	private static final long serialVersionUID = 5327837082727322924L;

	/**
	 * constructor
	 *
	 * @param throwable
	 * @param message
	 */
	public TechnicalException(Throwable throwable, String message) {
		super(message, throwable);
	}

	/**
	 * constructor
	 *
	 * @param throwable
	 */
	public TechnicalException(Throwable throwable) {
		super(throwable);
	}

	/**
	 * constructor
	 *
	 * @param message
	 * @param throwable
	 */
	public TechnicalException(String message) {
		super(message);
	}

}
