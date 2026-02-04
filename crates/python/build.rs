fn main() {
    // This is needed for PyO3 to work correctly
    pyo3_build_config::add_extension_module_link_args();
}
