<?php

/**
 * Main class for N-Order Sheet plugin
 *
 * @since 1.0.0
 */

defined( 'ABSPATH' ) || exit;

class NOrderSheet {

	/**
	 * The one and only true NOrderSheet instance
	 *
	 * @since 1.0.0
	 * @access private
	 * @var object $instance
	 */
	private static $instance;

	/**
	 * Plugin version
	 *
	 * @since 1.0.0
	 * @var string
	 */
	private $version = '1.0.0';

	/**
	 * Instantiate the main class
	 *
	 * This function instantiates the class, initialize all functions and return the object.
	 *
	 * @since 1.0.0
	 * @return object The one and only true NOrderSheet instance.
	 */
	public static function instance() {

		if ( ! isset( self::$instance ) && ( ! self::$instance instanceof NOrderSheet ) ) {

			self::$instance = new NOrderSheet;
			self::$instance->set_up_constants();
			self::$instance->includes();

		}

		return self::$instance;
	}

	/**
	 * Function for setting up constants
	 *
	 * This function is used to set up constants used throughout the plugin.
	 *
	 * @since 1.0.0
	 */
	public function set_up_constants() {

		self::set_up_constant( 'NORDER_SHEET_VERSION', $this->version );
		self::set_up_constant( 'NORDER_SHEET_PLUGIN_PATH', plugin_dir_path( __FILE__ ) . '../' );
		self::set_up_constant( 'NORDER_SHEET_PLUGIN_URL', plugin_dir_url( __FILE__ ) . '../' );
		self::set_up_constant( 'NORDER_SHEET_LIBRARIES_PATH', plugin_dir_path( __FILE__ ) );
		self::set_up_constant( 'NORDER_SHEET_DEBUG', true );

	}

	/**
	 * Make new constants
	 *
	 * @param string $name
	 * @param mixed $val
	 */
	public static function set_up_constant( $name, $val = false ) {

		if ( ! defined( $name ) ) {
			define( $name, $val );
		}

	}

	/**
	 * Includes all necessary PHP files
	 *
	 * This function is responsible for including all necessary PHP files.
	 *
	 * @since 1.0.0
	 */
	public function includes() {

		if ( defined( 'NORDER_SHEET_LIBRARIES_PATH' ) ) {
			require NORDER_SHEET_LIBRARIES_PATH . 'class-norder-sheet-api.php';
			require NORDER_SHEET_LIBRARIES_PATH . 'class-norder-sheet-settings.php';
			require NORDER_SHEET_LIBRARIES_PATH . 'class-norder-sheet-collector.php';
		}

	}

}